# PyC3Dserver
Python interface of C3Dserver software for reading and editing C3D motion capture files.

## Description
PyC3Dserver is a python inteface of C3Dserver using PyWin32.

## Installation
PyC3Dserver can be installed from [PyPI](https://pypi.org/project/pyc3dserver/) using ```pip``` on Python>=3.7.

```bash
pip install pyc3dserver
```

## Prerequisites
C3Dserver x64 edition (for Windows x64 platforms) installation: https://www.c3dserver.com/

## Usage
Most of numerical inputs and outputs for PyC3Dserver will be in the form of NumPy arrays. So let's import NumPy module.
```python
import numpy as np
```
After the installation of PyC3Dserver, you can import it as follows:
```python
import pyc3dserver as c3d
```
You can get the COM object of C3Dserver like this. You need to use this COM object while you are working with PyC3Dserver module.
```python
# Get the COM object of C3Dserver
itf = c3d.c3dserver()
```
Then, you can open a C3D file.
```python
# Open a C3D file
ret = c3d.open_c3d(itf, "sample_file.c3d")
```
Following functions are the most useful ones to extract the information from a C3D file. All the outputs are python dictionary types.
```python
# For the information of header
dict_header = c3d.get_dict_header(itf)
# For the information of all groups
dict_groups = c3d.get_dict_groups(itf)
# For the information of all markers(points)
dict_markers = c3d.get_dict_markers(itf)
# For the information of all forces/moments
dict_forces = c3d.get_dict_forces(itf)
# For the information of all analogs(excluding or including forces/moments)
dict_analogs = c3d.get_dict_analogs(itf)
```
If you made any modification in the C3Dserver and want to save it, you need to use the following function explicitly.
```python
# Save the C3D file from C3Dserver
ret = c3d.save_c3d(itf, "new_file.c3d")
```
After all your processes, it is recommended to close the C3D file from C3Dserver.
```python
# Close the C3D file from C3Dserver
ret = c3d.close_c3d(itf)
```

## Examples
There are more functions to get the information of individual markers and analogs. Also there are other functions for editing C3D files.
You can find some examples [here](https://github.com/mkjung99/pyc3dserver_examples).

## Limitations
PyC3Dserver tries to implement some useful functions using C3Dserver internally, but it does not cover full potential features of C3Dserver.
You can develop your own functions using the COM object of C3Dserver in Python.

## Dependencies
- PyWin32: ([GitHub](https://github.com/mhammond/pywin32), [PyPI](https://pypi.org/project/pywin32/), [Anaconda](https://anaconda.org/anaconda/pywin32))
- NumPy: ([Website](https://numpy.org/), [PyPI](https://pypi.org/project/numpy/), [Anaconda](https://anaconda.org/anaconda/numpy))
- SciPy: ([Website](https://www.scipy.org/), [PyPI](https://pypi.org/project/scipy/), [Anaconda](https://anaconda.org/anaconda/scipy))

## References
- [C3D.ORG](https://www.c3d.org/)
- [C3Dserver.com](https://www.c3dserver.com/)
- [Motion Lab Systems, Inc.](https://www.motion-labs.com/)
- [PyWin32](https://github.com/mhammond/pywin32)

## Python IDE recommendation
- [Spyder](https://www.spyder-ide.org/) for MATLAB friendly users
- [Visual Studio Code](https://code.visualstudio.com/)

## Acknowledgment
This work is funded by the European Union’s Horizon 2020 research and innovation programme (Project EXTEND - Bidirectional Hyper-Connected Neural System) under grant agreement No 779982.

## How to cite this work
[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.3908713.svg)](https://doi.org/10.5281/zenodo.3908713)

## License
[MIT](https://choosealicense.com/licenses/mit/)
