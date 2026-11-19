#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
python3 crear_paquete_portable.py "$@"
