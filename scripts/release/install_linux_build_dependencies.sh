#!/usr/bin/env bash
# Production builds use the workflow's immutable Ubuntu image and this archive snapshot.
set -euo pipefail

if [[ "$(id -u)" != 0 || ! -f /.dockerenv ]]; then
  echo 'Linux release dependencies must be installed inside the release container.' >&2
  exit 2
fi

source /etc/os-release
if [[ "$ID" != ubuntu || "$VERSION_ID" != 24.04 || "$(dpkg --print-architecture)" != amd64 ]]; then
  echo 'Linux release dependencies require Ubuntu 24.04 amd64.' >&2
  exit 2
fi

sources=$(mktemp /tmp/docwen-ubuntu-sources-XXXXXX.sources)
trap 'rm -f -- "$sources"' EXIT
cat > "$sources" <<'SOURCES'
Types: deb
URIs: https://snapshot.ubuntu.com/ubuntu/20260909T000000Z/
Suites: noble noble-updates noble-security
Components: main universe restricted multiverse
Signed-By: /usr/share/keyrings/ubuntu-archive-keyring.gpg
SOURCES

# Ignore additional image repositories so they cannot influence dependency selection.
apt_options=(-o "Dir::Etc::sourcelist=$sources" -o 'Dir::Etc::sourceparts=-')
apt-get "${apt_options[@]}" update -qq
DEBIAN_FRONTEND=noninteractive apt-get "${apt_options[@]}" install -y --no-install-recommends \
  binutils xvfb xauth libegl1 libgl1 libxkbcommon-x11-0 libxcb-cursor0 \
  libglib2.0-0t64 libfontconfig1 libnss3 libdbus-1-3 libgcrypt20 fonts-dejavu-core

# Retain the resolved package versions in the immutable workflow log.
dpkg-query -W -f='${binary:Package}\t${Version}\n' | LC_ALL=C sort
