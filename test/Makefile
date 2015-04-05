compiler = mcs
runtime = MONO_PATH=../build mono
src = GenodfTest.cs

exe = GenodfTest.exe

all: run

$(exe): $(src)
	mkdir -p build/
	$(compiler) $(src) -out:build/$(exe) -reference:../build/genodf.dll

run: $(exe)
	$(runtime) build/$(exe)

clean:
	rm -f build/*
