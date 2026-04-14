package main

import (
	_ "embed"
	"fmt"
	"log"
	"net"
	"net/http"
	"os/exec"
	"runtime"
	"time"
)

//go:embed pdf_merger.html
var pdfMergerHTML []byte

func openBrowser(url string) error {
	var cmd string
	var args []string
	switch runtime.GOOS {
	case "windows":
		cmd = "rundll32"
		args = []string{"url.dll,FileProtocolHandler", url}
	case "darwin":
		cmd = "open"
		args = []string{url}
	default: // linux, freebsd, etc.
		cmd = "xdg-open"
		args = []string{url}
	}
	return exec.Command(cmd, args...).Start()
}

func main() {
	mux := http.NewServeMux()
	mux.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/" && r.URL.Path != "/index.html" {
			http.NotFound(w, r)
			return
		}
		w.Header().Set("Content-Type", "text/html; charset=utf-8")
		w.Header().Set("Cache-Control", "no-store")
		_, _ = w.Write(pdfMergerHTML)
	})

	// Bind to a free port on localhost only.
	listener, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		log.Fatalf("Failed to start server: %v", err)
	}
	addr := listener.Addr().(*net.TCPAddr)
	url := fmt.Sprintf("http://127.0.0.1:%d/", addr.Port)

	fmt.Println("PDF Merger is running.")
	fmt.Printf("Open %s in your browser if it didn't open automatically.\n", url)
	fmt.Println("Close this window to stop the app.")

	// Give the server a moment to start, then open the browser.
	go func() {
		time.Sleep(300 * time.Millisecond)
		if err := openBrowser(url); err != nil {
			fmt.Printf("Could not open browser automatically: %v\n", err)
		}
	}()

	if err := http.Serve(listener, mux); err != nil {
		log.Fatalf("Server error: %v", err)
	}
}
