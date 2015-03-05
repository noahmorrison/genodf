name = genodf
compiler = mcs
src = src/*.cs

dll = $(name).dll

resourceNames = $(shell find resources/* -type f)
resources = $(foreach res, $(resourceNames), -resource:$(res))


all: $(dll)

$(dll): $(src)
	mkdir -p build/
	$(compiler) $(src) -out:build/$(dll) $(resources) -target:library

clean:
	rm -f build/$(dll)
