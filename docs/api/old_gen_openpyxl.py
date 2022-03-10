from pathlib import Path
import openpyxl

root = Path('openpyxl')
root.mkdir(exist_ok=True)
_root = Path(openpyxl.__file__).parent.resolve()
names = [file.name for file in _root.iterdir()
         if file.is_dir() and file.name != '__pycache__']
_names = '\n    '.join(names)
index = f'''openpyxl
=====================================

.. toctree::

    {_names}
'''

with open(root/'index.rst', 'w') as fp:
    fp.write(index)
for name in names:
    tree = []
    for file in (_root/name).iterdir():
        if file.name in ['__init__.py', '__pycache__']:
            continue
        else:
            if file.is_file():
                tree.append(file.name.replace('.py', ''))
    tree.append(name)
    tree = '\n    '.join(tree)
    doc = f'''openpyxl.{name}
===================================

.. currentmodule:: openpyxl.{name}

.. autosummary::
    :toctree: generated
    :nosignatures:
    :template: classtemplate.rst

    {tree}
'''
    with open(root/f'{name}.rst', 'w') as fp:
        fp.write(doc)
