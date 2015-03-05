name = genodf
compiler = mcs
src = OpenDocument.cs Resources.cs Spreadsheet.cs

dll = $(name).dll

resourceNames = $(shell find resources/* -type f)
resources = $(foreach res, $(resourceNames), -resource:$(res))


all: $(dll)

$(dll): $(src)
	mkdir -p build/
	$(compiler) $(src) -out:build/$(dll) $(resources) -target:library

clean:
	rm -f build/$(dll)
