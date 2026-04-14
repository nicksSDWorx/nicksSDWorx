#!/usr/bin/env bash
# Build the embedded Python + Tcl/Tk + pypdf + app bundle that
# pdf_merger.exe ships with.
#
# Outputs:   launcher/bundle.zip  (~10 MB)
#
# Requires:  curl, unzip, zstd, python3, pip, go
# Network:   conda.anaconda.org (for python + tk), pypi for pypdf

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BUILD="${HERE}/.build"
mkdir -p "${BUILD}"
cd "${BUILD}"

PY_CONDA_URL="https://conda.anaconda.org/conda-forge/win-64/python-3.11.15-h0159041_0_cpython.conda"
TK_CONDA_URL="https://conda.anaconda.org/conda-forge/win-64/tk-8.6.13-h5226925_1.conda"

echo "==> Downloading Windows Python (conda-forge)..."
[[ -f python.conda ]] || curl -sL -o python.conda "${PY_CONDA_URL}"
[[ -f tk.conda ]]     || curl -sL -o tk.conda     "${TK_CONDA_URL}"

echo "==> Extracting Python..."
rm -rf py-extract py-pkg
mkdir -p py-extract py-pkg
unzip -q python.conda -d py-extract/
zstd -qd py-extract/pkg-python-*.tar.zst -o py-pkg/pkg.tar
tar -xf py-pkg/pkg.tar -C py-pkg/
rm py-pkg/pkg.tar

echo "==> Extracting Tcl/Tk..."
rm -rf tk-extract tk-pkg
mkdir -p tk-extract tk-pkg
unzip -q tk.conda -d tk-extract/
zstd -qd tk-extract/pkg-tk-*.tar.zst -o tk-pkg/pkg.tar
tar -xf tk-pkg/pkg.tar -C tk-pkg/
rm tk-pkg/pkg.tar

echo "==> Downloading pypdf wheel..."
rm -rf pypdf-dl pypdf-ext
mkdir -p pypdf-dl pypdf-ext
pip3 download --no-deps --dest pypdf-dl --only-binary=:all: 'pypdf==6.10.0' >/dev/null
python3 -m zipfile -e pypdf-dl/pypdf-*.whl pypdf-ext/

echo "==> Assembling bundle..."
rm -rf bundle
mkdir -p bundle/Lib/site-packages bundle/tcl
cp py-pkg/python.exe     bundle/
cp py-pkg/pythonw.exe    bundle/
cp py-pkg/python3.dll    bundle/
cp py-pkg/python311.dll  bundle/
cp -r py-pkg/DLLs        bundle/
cp -r py-pkg/Lib/.       bundle/Lib/
cp tk-pkg/Library/bin/tcl86t.dll bundle/DLLs/
cp tk-pkg/Library/bin/tk86t.dll  bundle/DLLs/
cp tk-pkg/Library/bin/zlib1.dll  bundle/DLLs/ 2>/dev/null || true
cp -r tk-pkg/Library/lib/tcl8.6 bundle/tcl/
cp -r tk-pkg/Library/lib/tk8.6  bundle/tcl/
# Trim stdlib fat
rm -rf bundle/Lib/{idlelib,test,lib2to3,turtledemo,ensurepip,distutils,pydoc_data}
rm -rf bundle/Lib/unittest/test
find bundle -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
# Add pypdf + app
cp -r pypdf-ext/pypdf bundle/Lib/site-packages/
cp "${HERE}/pdf_merger.py" bundle/

echo "==> Packing bundle.zip..."
( cd bundle && zip -qr9 "${HERE}/launcher/bundle.zip" . )
ls -lh "${HERE}/launcher/bundle.zip"
echo "==> Done."
