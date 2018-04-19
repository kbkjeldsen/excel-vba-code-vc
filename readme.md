# Version control of VBA scripts
Tracking changes to VBA code in Excel files is done by using a python script to extract the VBA code of the Excel file into separate files as described here: https://www.xltrail.com/blog/auto-export-vba-commit-hook

The python script that extracts the VBA code (extract-vba.py) is placed directly in the repository root.

The script is called by use of a custom pre-commit [git hook](https://git-scm.com/docs/githooks) which is called before every commit.

## Dependencies
This example requires [python 3](https://www.python.org/downloads/windows/) and [Python Launcher](https://docs.python.org/3/using/windows.html#launcher)  to be installed on the local machine to work out-of-the-box.

## Change tracking of custom git hooks
By default, files in the ./git/hooks directory are not tracked and will not be a part of the project.

To be able to track changes, the custom hooks can be placed in another directory, which is being tracked and git can be instructed to use the hooks in this directory instead of the default hooks directory by adding the attribute core.hooksPath to the .git/config file. In this example, the custom hooks are placed in the "./custom_hooks" directory.

**NB!** *The git config file is **not** being tracked, so it must be updated whenever the repository has been cloned.*

## Custom pre-commit script
The pre-commit script in this example first executes the extract-vba.py script to extract the VBA code into separate directories. After that, a check is made for each of the possible Excel file extensions to see if any VBA code has been extracted, adn if so, the extracted files are added to the commit.