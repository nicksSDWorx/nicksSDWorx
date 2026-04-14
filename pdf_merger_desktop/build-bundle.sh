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

CF="https://conda.anaconda.org/conda-forge/win-64"
PY_CONDA_URL="${CF}/python-3.11.15-h0159041_0_cpython.conda"
TK_CONDA_URL="${CF}/tk-8.6.13-h5226925_1.conda"
# Runtime DLLs (MSVC + zlib + ssl + ffi + bzip2 + expat + lzma + sqlite)
RUNTIME_URLS=(
    "${CF}/vc14_runtime-14.42.34433-h6356254_33.conda"
    "${CF}/libzlib-1.3.1-h2466b09_2.conda"
    "${CF}/openssl-3.6.2-hf411b9b_0.conda"
    "${CF}/libffi-3.5.2-h52bdfb6_0.conda"
    "${CF}/bzip2-1.0.8-hcfcfb64_5.conda"
    "${CF}/libexpat-2.7.5-hac47afa_0.conda"
    "${CF}/liblzma-5.8.3-hfd05255_0.conda"
    "${CF}/libsqlite-3.53.0-hf5d6505_0.conda"
)

fetch_conda_pkg() {
    local url="$1" out="$2"
    local archive="${out}.conda"
    [[ -f "$archive" ]] || curl -sL -o "$archive" "$url"
    rm -rf "${out}-e" "${out}-p"
    mkdir -p "${out}-e" "${out}-p"
    unzip -q "$archive" -d "${out}-e/"
    zstd -qd "${out}-e"/pkg-*.tar.zst -o "${out}-p/pkg.tar"
    tar -xf "${out}-p/pkg.tar" -C "${out}-p/"
    rm "${out}-p/pkg.tar"
}

echo "==> Downloading Windows Python + Tcl/Tk + runtime DLLs (conda-forge)..."
fetch_conda_pkg "${PY_CONDA_URL}" "py"
fetch_conda_pkg "${TK_CONDA_URL}" "tk"

# Fetch all runtime DLL packages
mkdir -p runtime-dlls
for url in "${RUNTIME_URLS[@]}"; do
    name="${url##*/}"                 # foo-1.2.3-abc.conda
    name="${name%.conda}"             # foo-1.2.3-abc
    slug="${name%%-*}"                # foo
    fetch_conda_pkg "$url" "rt-$slug"
    # Each package puts its DLLs under Library/bin or at the root
    find "rt-${slug}-p" \( -name "*.dll" \) -exec cp -u {} runtime-dlls/ \;
done
ls runtime-dlls/

echo "==> Downloading pypdf wheel..."
rm -rf pypdf-dl pypdf-ext
mkdir -p pypdf-dl pypdf-ext
pip3 download --no-deps --dest pypdf-dl --only-binary=:all: 'pypdf==6.10.0' >/dev/null
python3 -m zipfile -e pypdf-dl/pypdf-*.whl pypdf-ext/

echo "==> Assembling bundle..."
rm -rf bundle
mkdir -p bundle/Lib/site-packages bundle/tcl
cp py-p/python.exe     bundle/
cp py-p/pythonw.exe    bundle/
cp py-p/python3.dll    bundle/
cp py-p/python311.dll  bundle/
cp -r py-p/DLLs        bundle/
cp -r py-p/Lib/.       bundle/Lib/
cp tk-p/Library/bin/tcl86t.dll bundle/DLLs/
cp tk-p/Library/bin/tk86t.dll  bundle/DLLs/
cp -r tk-p/Library/lib/tcl8.6 bundle/tcl/
cp -r tk-p/Library/lib/tk8.6  bundle/tcl/
# Runtime DLLs — placed in both the Python dir and DLLs/ so Windows'
# DLL search finds them when loading python311.dll and every .pyd.
cp runtime-dlls/*.dll bundle/
cp runtime-dlls/*.dll bundle/DLLs/
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
