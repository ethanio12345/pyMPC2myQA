# README.md

Placeholder for when I write something later


## Build instruction
pyinstaller needs path to top level of project rather than subfolder that holds the packages (i.e. top level or project or top level with __init__)
Also, pyinstaller needs to be run from .venv if not wanting to add absolutely every single module in existence from absolutely everything into build.

So, command for single file mode is;
	pyinstaller ReadandConvertMPCFolder.py --clean -y -F --paths /path/to/pyMPC2myQA/