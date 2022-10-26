from setuptools import setup

setup(
    name='documentPro',
    version='0.0.2',
    entry_points={
        'console_scripts': [
            'docpro=utils:run',
            'docproapp=app:main'
        ]
    },
    install_requires=[
        'pywin32',
    ],
    py_modules=[
        'utils',
        'app'
    ]
)