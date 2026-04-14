//go:build windows

package main

import (
	"os/exec"
	"syscall"
	"unsafe"
)

func hidePythonConsole(cmd *exec.Cmd) {
	cmd.SysProcAttr = &syscall.SysProcAttr{
		HideWindow:    true,
		CreationFlags: 0x08000000, // CREATE_NO_WINDOW
	}
}

// messageBox shows a Windows MessageBox via user32.dll (no cgo).
func messageBox(title, text string) error {
	user32 := syscall.NewLazyDLL("user32.dll")
	proc := user32.NewProc("MessageBoxW")
	t, err := syscall.UTF16PtrFromString(text)
	if err != nil {
		return err
	}
	cap, err := syscall.UTF16PtrFromString(title)
	if err != nil {
		return err
	}
	const MB_ICONERROR = 0x00000010
	const MB_OK = 0x00000000
	_, _, _ = proc.Call(0, uintptr(unsafe.Pointer(t)), uintptr(unsafe.Pointer(cap)), uintptr(MB_ICONERROR|MB_OK))
	return nil
}
