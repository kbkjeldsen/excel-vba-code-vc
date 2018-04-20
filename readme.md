# Version control of VBA scripts
This repo hosts an example of how to use a Python script and pre-commit hooks of git to track changes to VBA code in Excel files.

The change tracking of VBA code is achieved by using a python script to extract the VBA code of the Excel file into separate files. It is based on the following article by Björn Stiel: https://www.xltrail.com/blog/auto-export-vba-commit-hook

## Dependencies
This example requires [python 3](https://www.python.org/downloads/windows/) and [Python Launcher](https://docs.python.org/3/using/windows.html#launcher) to be installed on the local machine to work out-of-the-box. The extract-vba.py script has a dependency against [oletools](https://github.com/decalage2/oletools) which can be installed by use of Python's package manager: `pip install -U oletools`

## How it works
The python script `extract-vba.py` which is used to extract the VBA code, is placed directly in the repository root. The script extracts any VBA code found in any Excel file into separate files which are placed in a subdirectory named the same as the respective Excel file including extension and suffixed with '.vba'.

The script is called by use of a custom pre-commit [git hook](https://git-scm.com/docs/githooks) which is called before every commit as described below.

## Change tracking of custom git hooks
By default, files in the ./git/hooks directory are not tracked and will not be a part of the project.

To be able to track changes, the custom hooks can be placed in another directory, which is being tracked and git can be instructed to use the hooks in this directory instead of the default hooks directory by adding the attribute `core.hooksPath` to the `.git/config` file. In this example, the custom hooks are placed in the `./custom_hooks` directory.

**NB!** *The git config file is **not** being tracked since it is located inside the untracked `.git` directory, so the `core.hooksPath` must be added whenever the repository has been cloned.*

## Custom pre-commit script
The pre-commit script in this example first executes the extract-vba.py script to extract the VBA code into separate directories. After that, a check is made for each of the possible Excel file extensions to see if any VBA code has been extracted, adn if so, the extracted files are added to the commit.
