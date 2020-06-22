# PyC3Dserver
Python interface of C3Dserver software for reading and editing C3D motion capture files.

## Description
PyC3Dserver is a python interace of C3Dserver using PyWin32.

## Installation
PyC3Dserver can be installed from [PyPI](https://pypi.org/project/pyc3dserver/) using ```pip``` on Python>=3.7.

```bash
pip install pyc3dserver
```

## Prerequisites
C3Dserver x64 edition (for Windows x64 platforms) installation: https://www.c3dserver.com/

## Usage
Most of numerial inputs will be in form of numpy arrays. So it is recommended to import numpy as well.
```python
import numpy as np
```
After the installation of PyC3Dserver, you can import it as follows:
```python
import pyc3dserver as c3d
```
You can get the COM object of C3Dserver like this. You need to use this COM object while you are working with PyC3Dserver module.
```python
itf = c3d.c3dserver()
```
Then, you can open a C3D file.
```python
ret = c3d.open_c3d(itf, "sample.c3d")
```
Following functions are most useful in order extract the information from a C3D file. All the outputs are python dictionary types.
```python
dict_header = c3d.get_dict_header(itf)
dict_groups = c3d.get_dict_groups(itf)
dict_markers = c3d.get_dict_markers(itf)
dict_forces = c3d.get_dict_forces(itf)
dict_analogs = c3d.get_dict_analogs(itf)
```
After all your processes, it is recommended to close the open C3D file from C3Dserver.
```python
ret = c3d.close_c3d(itf)
```


## Examples
Find the [examples](https://github.com/mkjung99/pyc3dserver_examples).

## Dependencies
- PyWin32: ([GitHub](https://github.com/mhammond/pywin32), [PyPI](https://pypi.org/project/pywin32/), [Anaconda](https://anaconda.org/anaconda/pywin32))
- NumPy: ([Website](https://numpy.org/), [PyPI](https://pypi.org/project/numpy/), [Anaconda](https://anaconda.org/anaconda/numpy))
- SciPy: ([Website](https://www.scipy.org/), [PyPI](https://pypi.org/project/scipy/), [Anaconda](https://anaconda.org/anaconda/scipy))

## References
- [C3D.ORG](https://www.c3d.org/)
- [C3Dserver.com](https://www.c3dserver.com/)
- [Motion Lab Systems, Inc.](https://www.motion-labs.com/)
- [PyWin32](https://github.com/mhammond/pywin32)
