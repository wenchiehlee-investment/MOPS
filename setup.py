"""Setup configuration for MOPS downloader package."""

from setuptools import setup, find_packages
from pathlib import Path

# Read README
readme_file = Path(__file__).parent / "README.md"
long_description = readme_file.read_text(encoding='utf-8') if readme_file.exists() else ""

# Read requirements
requirements_file = Path(__file__).parent / "requirements.txt"
if requirements_file.exists():
    requirements = requirements_file.read_text().strip().split('\n')
    requirements = [req.strip() for req in requirements if req.strip() and not req.startswith('#')]
else:
    requirements = [
        'requests>=2.28.0',
        'beautifulsoup4>=4.11.0',
        'lxml>=4.9.0',
    ]

setup(
    name="mops-downloader",
    version="1.0.0",
    author="MOPS Downloader Team",
    description="Download Taiwan MOPS quarterly financial reports",
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=find_packages(),
    python_requires=">=3.9",
    install_requires=requirements,
    entry_points={
        'console_scripts': [
            'mops-downloader=mops_downloader.cli:main',
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Financial and Insurance Industry",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business :: Financial",
        "Topic :: Internet :: WWW/HTTP :: Dynamic Content",
    ],
    keywords="taiwan, mops, financial, reports, downloader, ifrs",
    project_urls={
        "Bug Reports": "https://github.com/your-repo/mops-downloader/issues",
        "Source": "https://github.com/your-repo/mops-downloader",
    },
)
