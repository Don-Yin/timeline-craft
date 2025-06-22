from setuptools import setup, find_packages
from pathlib import Path

# Read the contents of README file
readme_path = Path(__file__).parent / "README.MD"
long_description = readme_path.read_text(encoding="utf-8") if readme_path.exists() else ""

# Read requirements from requirements.txt
requirements_path = Path(__file__).parent / "requirements.txt"
install_requires = []
if requirements_path.exists():
    install_requires = requirements_path.read_text(encoding="utf-8").strip().split("\n")

setup(
    name="timeline-craft",
    version="0.0.2",
    packages=find_packages(),
    install_requires=install_requires,
    author="Don Yin",
    author_email="don_yin@outlook.com",
    description="A Python library for creating timeline visualizations in PowerPoint",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Don-Yin/timeline-craft",
    python_requires=">=3.12",
    include_package_data=True,
)
