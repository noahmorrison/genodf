name = genodf
compiler = mcs
src = $(shell find src/* -type f)

dll = $(name).dll

resourceNames = $(shell find resources/* -type f)
resources = $(foreach res, $(resourceNames), -resource:$(res))

referenceNames = $(shell find references/* -type f)
references = $(foreach ref, $(referenceNames), -reference:$(ref))


all: $(dll)

$(dll): $(src)
	mkdir -p build/
	$(compiler) $(src) -out:build/$(dll) $(resources) $(references) -target:library
	cp -r references/* build/

clean:
	rm -f build/$(dll)
