from setuptools import setup, find_namespace_packages

setup(
    name="cli-anything-pages",
    version="1.0.0",
    description="CLI harness for Apple Pages — agent-operable document creation",
    author="Marcel R. G. Berger",
    author_email="hello@marcelrgberger.com",
    packages=find_namespace_packages(include=["cli_anything.*"]),
    install_requires=[
        "click>=8.0.0",
        "prompt-toolkit>=3.0.0",
    ],
    entry_points={
        "console_scripts": [
            "cli-anything-pages=cli_anything.pages.pages_cli:main",
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
