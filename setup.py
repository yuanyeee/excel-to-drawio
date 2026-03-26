#!/usr/bin/env python3
"""
Setup script for excel-to-drawio
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as f:
    long_description = f.read()

setup(
    name="excel-to-drawio",
    version="0.1.0",
    author="yuanyeee",
    author_email="yuanyeee@gmail.com",
    description="Convert Excel shapes/diagrams to draw.io format",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yuanyeee/excel-to-drawio",
    packages=find_packages(),
    install_requires=[
        "openpyxl>=3.1.0",
        "click>=8.1.0",
    ],
    entry_points={
        "console_scripts": [
            "excel-to-drawio=main:main",
        ],
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Topic :: Office/Business",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
    ],
    python_requires=">=3.8",
)
