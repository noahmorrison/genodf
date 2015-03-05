compiler = mcs
runtime = MONO_PATH=../build mono
src = GenodfTest.cs

exe = GenodfTest.exe

all: run

$(exe): $(src)
	$(compiler) $(src) -out:$(exe) -reference:../build/genodf.dll

run: $(exe)
	$(runtime) $(exe)

clean:
	rm -f $(exe)
