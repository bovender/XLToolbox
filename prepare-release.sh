#!/bin/bash
# Helper script to prepare an XL Toolbox NG release.
# (c) Daniel Kraus 2015

# As a precaution, do not do anything if there are uncommitted changes
# in the Git repository.
if ! git diff-index HEAD --quiet --; then
        echo "There are uncommitted changes."
        echo "Please call this script with a clean working directory."
        exit 1
fi

# Get current git branch
# From http://stackoverflow.com/a/1593487/270712
BRANCH="$(git symbolic-ref HEAD --short 2>/dev/null)"
echo "Current git branch: $BRANCH"

# Extract the new version string from the branch name,
# or take the first argument as version string if the
# current branch is not a release branch.
if [[ "$BRANCH" != release-* ]]; then
        if [ $# -eq 0 ]; then
                echo "Please create a release branch named 'release-[SEMANTIC VERSION]',"
                echo "or call this script with an argument to indicate the new SEMANTIC VERSION."
                exit 1
        fi
        VERSION="$1"
else
        VERSION=${BRANCH#release-}
fi

# Find out if a tag with this version string exists already
TAG="v$VERSION"
if git rev-parse "$TAG" >/dev/null 2>&1; then
        echo "Found a tag $TAG in this repository."
        echo "It appears that $VERSION has been released already."
        exit 2
fi

# Check out a release branch if we are not on one already
if [ "$BRANCH" != release-* ]; then
        RELEASEBRANCH="release-$VERSION"
        # Could use -B option to have git automatically create the branch
        # if it does not exist, but this will also reset any existing
        # release branch, which may be an unwanted side effect.
        git checkout "$RELEASEBRANCH" 2>/dev/null || git checkout -b "$RELEASEBRANCH" || exit 3
fi

# Now for the real work...
echo "Preparing new version $VERSION"
MAJORMINORPATCH=${VERSION%%-*}
MSVERSION=$(tail -n 1 VERSION)
if [[ "$MSVERSION" == "$MAJORMINORPATCH"* ]]; then
        INC=${MSVERSION##*.}
        ((INC++))
        MSVERSION="${MAJORMINORPATCH}.${INC}"
else
        MSVERSION="${MAJORMINORPATCH}.0"
fi

echo "Writing new VERSION file:"
echo $VERSION > VERSION
echo $MSVERSION >> VERSION
cat VERSION

# Update MS version numbers in AssemblyInfo files
find . -maxdepth 3 -name 'AssemblyInfo.cs' -execdir \
        sed -i -r -e 's/(^\[assembly: Assembly(File)?Version\(")[^"]+(.*)$/\1'$MSVERSION'\3/' \{} \;

echo "Finished."
echo "Examine the changed files, then commit them on this branch."
git status
