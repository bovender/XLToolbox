SHELL := /bin/bash
.PHONY: credits publish all

help:
	# Target publish-alpha: Pushes the current alpha version to the server using the latest tag
	#                       Also pushes the branch and tags to the remote repository

credits: XLToolbox/Resources/html/credits.html

XLToolbox/Resources/html/credits.html: ../web/content/about.haml
	sed -e '1,/<!-- 8< -->/d; /vim:/d; s/^\( \)\{4\}//' ../web/content/about.haml | perl -0777 -pe 's/\[([^]]+)\]\([^)]+\)/\1/msg' | pandoc -H XLToolbox/Resources/html/style.html > XLToolbox/Resources/html/credits.html

publish:
	git push
	git push --tags
	publish/create-release.rb
update-bovender:
	find XLToolbox XLToolboxForExcel Tests -name '*.csproj' -o -name 'packages.config' -print0 | xargs -0 sed -i 's/0\.16\.1/0.16.2/'
