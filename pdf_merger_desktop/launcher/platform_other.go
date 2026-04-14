//go:build !windows

package main

import (
	"fmt"
	"os/exec"
)

func hidePythonConsole(cmd *exec.Cmd) {}

func messageBox(title, text string) error {
	return fmt.Errorf("messageBox not supported on this platform")
}
