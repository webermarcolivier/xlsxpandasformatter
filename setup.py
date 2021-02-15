from setuptools import setup, find_packages

with open("README.md", encoding="utf-8") as f:
    readme = f.read()

about = {}
with open("xlsxpandasformatter/_version.py", encoding="utf-8") as f:
    exec(f.read(), about)

setup(
    name="XlsxPandasFormatter",
    version=about["__version__"],
    description="XlsxPandasFormatter",
    long_description=readme,
    long_description_content_type="text/markdown",
    url="https://github.com/webermarcolivier/xlsxpandasformatter",
    author="Marc Weber",
    author_email="329640305@qq.com",
    license="MIT",
    packages=find_packages(),
    zip_safe=False,
    python_requires=">=3.6",
    install_requires=["pandas", "XlsxWriter", "matplotlib", "numpy", "seaborn"],
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3 :: Only",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
    ],
    keywords="xlsxpandasformatter xlsx excel pandas formatter",
    project_urls={
        "Repository": "https://github.com/webermarcolivier/xlsxpandasformatter",
        "Bug Reports": "https://github.com/webermarcolivier/xlsxpandasformatter/issues",
        "Documentation": "https://github.com/webermarcolivier/xlsxpandasformatter/blob/master/README.md",
    }
)
