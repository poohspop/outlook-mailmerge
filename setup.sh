#!/bin/bash
# setup.sh - Configure manifest with your GitHub username
if [ -z "$1" ]; then
  echo "Usage: ./setup.sh YOUR_GITHUB_USERNAME"
  exit 1
fi
sed -i "s/YOURGHUSER/$1/g" manifest.xml
echo "manifest.xml updated for https://$1.github.io/outlook-mailmerge/"
