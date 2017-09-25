# MDT Install SPP Script

## Description

This script helps to install HP Support Pack for Proliant (SPP) with MDT.

## How to use it

- Create a MDT application with all this script content (.ps1 + Sources folder)
- As command line, specify the InstallSPP.ps1 file
- Under the **Sources** folder, copy the content of hp\swpackages folder (content of the SPP ISO file)
  - You can remove all *.rpm items to reduce folder size (Linux packages)
