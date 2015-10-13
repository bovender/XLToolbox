SHELL := /bin/bash
.PHONY: credits publish-alpha all

help:
	# Target publish-alpha: Pushes the current alpha version to the server using the latest tag
	#                       Also pushes the branch and tags to the remote repository

credits: XLToolbox/html/credits.html

XLToolbox/html/credits.html: ../web/content/about.html.haml
	sed -e '1,/<!-- 8< -->/d; /vim:/d; s/^\(  \)\{3\}//' ../web/content/about.html.haml | perl -0777 -pe 's/\[([^]]+)\]\([^)]+\)/\1/msg' | pandoc -H XLToolbox/html/style.html > XLToolbox/html/credits.html

GITTAG=$(shell git describe master)
VERSION=$(GITTAG:v%=%)
publish-alpha:
	scp publish/release/XL_Toolbox_$(VERSION).exe bovender@frs.sourceforge.net:/home/frs/project/xltoolbox/alpha/
	git push
	git push --tags
