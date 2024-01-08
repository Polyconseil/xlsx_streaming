Releasing a new version
=======================

Installation
------------

Install Zest tool that automates process of releasing a new version of a Python package::

    $ pip install zest.releaser

Clean master branch
-------------------

Make sure you are working on a safe master branch::

    $ git checkout master
    $ git reset origin/master --hard
    $ make clean

Prerelease
----------

Once you are ready to release a new version, let's begin with the first command
``prerelease``. This command aims to change the setup.cfg's version and add the
current date to the changelog.

If you agree with the prefilled version, you can directly run::

    $ prerelease --no-input

Otherwise, run the command without option. Zest will ask you a version number. Then
valid commit changes.

Release
-------

Then, let's focus on the ``release`` command. This command focuses on creating a tag and
uploading the new package to PyPI.

The command line asks you to valid the tag command. Press ``Y``

The next question is about uploading the package to a custom repository. Press ``n``.

Postrelease
-----------

Finally, it's time to prepare the project back to development and push commits remotely.

Run the command::

    $ postrelease --feature

The first question is about the next development version. By default, the minor number
will be increased. If you want to enter another version, please do not add ``.dev0`` as
it is automatically added by Zest.

Then, valid commit changes. Press ``Y``.

Finally, you need to push local commits on the server. Press ``Y``.

Upload of the new release to PyPI
--------------------------------

This section is purely informative, everything is done automatically. Integration to
PyPI is done through the use of Github CI/CD. Declaration of the pipeline is done in
``.github/workflows/publish-to-pipy.yml``

On tag push, the job `publish-to-pypi` is triggered. The pipeline is composed of 3 main
steps:

* Setup Python and install build dependency
* Build the new project release
* Upload artifacts to PyPI

Setup of this pipeline is inspired from this `guide`_.

.. _guide: https://packaging.python.org/en/latest/guides/publishing-package-distribution-releases-using-github-actions-ci-cd-workflows/
