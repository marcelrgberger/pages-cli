from setuptools import setup, find_packages

setup(
    name="pages-cli",
    version="1.0.0",
    description="CLI for Apple Pages — create, edit, format, and export documents",
    author="Marcel R. G. Berger",
    author_email="hello@marcelrgberger.com",
    packages=find_packages(),
    install_requires=[
        "click>=8.0.0",
        "prompt-toolkit>=3.0.0",
    ],
    entry_points={
        "console_scripts": [
            "pages-cli=pages_cli.pages_cli:main",
        ],
    },
    python_requires=">=3.10",
    classifiers=[
        "Development Status :: 4 - Beta",
        "Environment :: Console",
        "Intended Audience :: Developers",
        "Operating System :: MacOS",
        "Programming Language :: Python :: 3",
        "Topic :: Office/Business :: Office Suites",
    ],
)
