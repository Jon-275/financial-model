#!/usr/bin/env bash
set -euo pipefail

# Calc comes from Ubuntu. Runner image feeds for other software are irrelevant here.
ubuntu_sources=(
  -o Dir::Etc::sourcelist=/etc/apt/sources.list.d/ubuntu.sources
  -o Dir::Etc::sourceparts=-
)
sudo apt-get "${ubuntu_sources[@]}" update --error-on=any
sudo apt-get "${ubuntu_sources[@]}" install -y --no-install-recommends libreoffice-calc
soffice --version
