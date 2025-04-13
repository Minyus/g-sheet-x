from setuptools import find_packages, setup

console_scripts = [
    "gsheetx = gsheetx.__main__:main",
]


with open("requirements.txt", encoding="utf-8") as f:
    requires = []
    for line in f:
        req = line.split("#", 1)[0].strip()
        if req and not req.startswith("--"):
            requires.append(req)

long_description = """
Please see:
https://github.com/Minyus/g-sheet-x
"""

setup(
    name="gsheetx",
    version="0.0.2",
    packages=find_packages(exclude=["tests"]),
    entry_points={"console_scripts": console_scripts},
    install_requires=requires,
    description="CLI for Google Spreadsheet",
    license="MIT",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Minyus/g-sheet-x",
    author="Yusuke Minami",
    author_email="me@minyus.github.com",
    zip_safe=False,
    keywords=", ".join(["Google", "sheet", "gspread"]),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Natural Language :: English",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Operating System :: OS Independent",
    ],
)
