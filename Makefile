.PHONY: credits 

credits: XLToolbox/html/credits.html

XLToolbox/html/credits.html: ../web/content/about.html.haml
	sed -e '1,/<!-- 8< -->/d; /vim:/d; s/^\(  \)\{3\}//' ../web/content/about.html.haml | perl -0777 -pe 's/\[([^]]+)\]\([^)]+\)/\1/msg' | pandoc -H XLToolbox/html/style.html > XLToolbox/html/credits.html


