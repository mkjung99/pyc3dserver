import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="pyc3dserver",
    version="0.1.2",
    author="Moon Ki Jung (et al.)",
    author_email="m.k.jung@outlook.com",
    description="Python interface of C3Dserver software for reading and editing C3D motion capture files.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/mkjung99/pyc3dserver",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows :: Windows 10",
        'Topic :: Scientific/Engineering :: Information Analysis',
        'Intended Audience :: Science/Research',
    ],
    python_requires='>=3.7',
    install_requires=['numpy', 'scipy', 'pywin32>=223'],
)