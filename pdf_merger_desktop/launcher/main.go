// pdf_merger.exe — native Windows launcher for the PDF Merger desktop app.
//
// The launcher embeds a minimal Windows Python runtime + the app's Python
// source + pypdf. On first run it extracts the bundle into
// %LOCALAPPDATA%\PdfMerger\<version-hash>\ and then launches
// pythonw.exe pdf_merger.py. No browser, fully offline.
package main

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	_ "embed"
	"fmt"
	"io"
	"io/fs"
	"os"
	"os/exec"
	"path/filepath"
	"time"
)

//go:embed bundle.zip
var bundleZip []byte

const appDirName = "PdfMerger"

func main() {
	if err := run(); err != nil {
		showError(err.Error())
		os.Exit(1)
	}
}

func run() error {
	// Version key derived from the embedded bundle so upgrades re-extract.
	sum := sha256.Sum256(bundleZip)
	versionKey := hex.EncodeToString(sum[:])[:16]

	baseDir, err := appDataDir()
	if err != nil {
		return fmt.Errorf("cannot determine app data dir: %w", err)
	}
	installDir := filepath.Join(baseDir, appDirName, versionKey)
	readyMarker := filepath.Join(installDir, ".ready")

	if _, err := os.Stat(readyMarker); err != nil {
		// Not installed — extract the bundle.
		if err := os.MkdirAll(installDir, 0o755); err != nil {
			return fmt.Errorf("create install dir: %w", err)
		}
		if err := extractZip(bundleZip, installDir); err != nil {
			return fmt.Errorf("extract bundle: %w", err)
		}
		if err := os.WriteFile(readyMarker, []byte("ok"), 0o644); err != nil {
			return fmt.Errorf("write ready marker: %w", err)
		}
		// Best-effort: clean up older versions.
		cleanOldVersions(filepath.Join(baseDir, appDirName), versionKey)
	}

	pythonw := filepath.Join(installDir, "pythonw.exe")
	script := filepath.Join(installDir, "pdf_merger.py")
	if _, err := os.Stat(pythonw); err != nil {
		return fmt.Errorf("pythonw.exe not found in install dir: %w", err)
	}
	if _, err := os.Stat(script); err != nil {
		return fmt.Errorf("pdf_merger.py not found in install dir: %w", err)
	}

	cmd := exec.Command(pythonw, script)
	cmd.Dir = installDir
	// Isolate Python so it only sees our bundled Lib and site-packages.
	env := cleanPythonEnv(os.Environ())
	env = append(env, "PYTHONHOME="+installDir)
	env = append(env, "PYTHONPATH="+filepath.Join(installDir, "Lib")+string(os.PathListSeparator)+filepath.Join(installDir, "Lib", "site-packages"))
	env = append(env, "TCL_LIBRARY="+filepath.Join(installDir, "tcl", "tcl8.6"))
	env = append(env, "TK_LIBRARY="+filepath.Join(installDir, "tcl", "tk8.6"))
	env = append(env, "PYTHONDONTWRITEBYTECODE=1")
	env = append(env, "PYTHONNOUSERSITE=1")
	// Prepend the install dir to PATH so dependent DLLs (zlib, openssl, etc.)
	// resolve from the bundle even if the system has different versions.
	env = append(env, prependPath(os.Getenv("PATH"), installDir, filepath.Join(installDir, "DLLs")))
	cmd.Env = env

	// Capture startup output to a log file so we can diagnose early crashes.
	logPath := filepath.Join(installDir, "launch.log")
	logFile, logErr := os.Create(logPath)
	if logErr == nil {
		cmd.Stdout = logFile
		cmd.Stderr = logFile
		defer logFile.Close()
	}

	hidePythonConsole(cmd)
	if err := cmd.Start(); err != nil {
		return fmt.Errorf("launch python: %w", err)
	}

	// Wait briefly to catch immediate startup failures (missing DLLs,
	// PYTHONHOME wrong, etc.). Tkinter's mainloop keeps the process alive
	// forever once it starts, so any early exit means something broke.
	errCh := make(chan error, 1)
	go func() { errCh <- cmd.Wait() }()
	select {
	case err := <-errCh:
		// Python exited within the grace period — show the tail of the log.
		tail, _ := readTail(logPath, 4000)
		msg := "Python exited immediately."
		if err != nil {
			msg += "\n\nExit status: " + err.Error()
		}
		if tail != "" {
			msg += "\n\nOutput:\n" + tail
		} else {
			msg += "\n\nLog file: " + logPath
		}
		return fmt.Errorf("%s", msg)
	case <-time.After(1500 * time.Millisecond):
		// Still running — assume the GUI is up. Detach.
		return nil
	}
}

func prependPath(existing string, dirs ...string) string {
	sep := string(os.PathListSeparator)
	joined := ""
	for _, d := range dirs {
		if joined != "" {
			joined += sep
		}
		joined += d
	}
	if existing != "" {
		joined += sep + existing
	}
	return "PATH=" + joined
}

func readTail(path string, maxBytes int64) (string, error) {
	f, err := os.Open(path)
	if err != nil {
		return "", err
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		return "", err
	}
	start := info.Size() - maxBytes
	if start < 0 {
		start = 0
	}
	if _, err := f.Seek(start, 0); err != nil {
		return "", err
	}
	b, err := io.ReadAll(f)
	if err != nil {
		return "", err
	}
	return string(b), nil
}

func cleanPythonEnv(env []string) []string {
	out := make([]string, 0, len(env))
	for _, kv := range env {
		switch {
		case startsWithI(kv, "PYTHONHOME="),
			startsWithI(kv, "PYTHONPATH="),
			startsWithI(kv, "PYTHONSTARTUP="),
			startsWithI(kv, "PYTHONUSERBASE="),
			startsWithI(kv, "TCL_LIBRARY="),
			startsWithI(kv, "TK_LIBRARY="),
			startsWithI(kv, "PATH="):
			continue
		}
		out = append(out, kv)
	}
	return out
}

func startsWithI(s, prefix string) bool {
	if len(s) < len(prefix) {
		return false
	}
	for i := 0; i < len(prefix); i++ {
		a, b := s[i], prefix[i]
		if a >= 'a' && a <= 'z' {
			a -= 32
		}
		if b >= 'a' && b <= 'z' {
			b -= 32
		}
		if a != b {
			return false
		}
	}
	return true
}

func appDataDir() (string, error) {
	if d := os.Getenv("LOCALAPPDATA"); d != "" {
		return d, nil
	}
	if d := os.Getenv("APPDATA"); d != "" {
		return d, nil
	}
	return os.UserCacheDir()
}

func cleanOldVersions(appRoot, keep string) {
	entries, err := os.ReadDir(appRoot)
	if err != nil {
		return
	}
	for _, e := range entries {
		if e.IsDir() && e.Name() != keep {
			_ = os.RemoveAll(filepath.Join(appRoot, e.Name()))
		}
	}
}

func extractZip(data []byte, dest string) error {
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return err
	}
	for _, f := range r.File {
		target := filepath.Join(dest, f.Name)
		// Zip-slip guard.
		rel, err := filepath.Rel(dest, target)
		if err != nil || rel == ".." || hasDotDotPrefix(rel) {
			return fmt.Errorf("unsafe path in zip: %q", f.Name)
		}
		if f.FileInfo().IsDir() {
			if err := os.MkdirAll(target, 0o755); err != nil {
				return err
			}
			continue
		}
		if err := os.MkdirAll(filepath.Dir(target), 0o755); err != nil {
			return err
		}
		if err := writeZipFile(f, target); err != nil {
			return err
		}
	}
	return nil
}

func hasDotDotPrefix(p string) bool {
	return len(p) >= 3 && p[0] == '.' && p[1] == '.' && (p[2] == '/' || p[2] == '\\')
}

func writeZipFile(f *zip.File, target string) error {
	rc, err := f.Open()
	if err != nil {
		return err
	}
	defer rc.Close()
	mode := f.Mode()
	if mode == 0 {
		mode = 0o644
	}
	out, err := os.OpenFile(target, os.O_CREATE|os.O_WRONLY|os.O_TRUNC, mode|fs.FileMode(0o200))
	if err != nil {
		return err
	}
	if _, err := io.Copy(out, rc); err != nil {
		out.Close()
		return err
	}
	return out.Close()
}

func showError(msg string) {
	if err := messageBox("PDF Merger", msg); err != nil {
		fmt.Fprintln(os.Stderr, "PDF Merger error:", msg)
	}
}
