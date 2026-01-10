# Argentum Online Localindex.dat Creator

This tool generates the `localindex.dat` file, which contains client-side texts used to reduce server bandwidth usage.

Original reference:
https://github.com/ao-org/Recursos/blob/master/init/localindex.dat

![image](https://github.com/ao-org/argentum20-creador-indices/assets/5874806/f143873f-8247-4755-a084-9f149ac2cac4)

---

## Usage

### Executable (recommended)

To generate the `localindex.dat` file, go to the `argentum-online-creador-indices` folder and run:


The executable must be executed from within the project directory so it can correctly access the required files.

---

### Development / Updating the index generator

If new entries or logic need to be added to the index generator:

1. Edit the script:

2. Regenerate the executable using PyInstaller:

3. Replace the existing `generar_localindex.exe` with the newly generated version.

---

## Notes

- This repository contains the **source code** (`.py`) used to generate `localindex.dat`.
- The executable (`.exe`) is provided for convenience and should be regenerated whenever changes are made to the generator script.

