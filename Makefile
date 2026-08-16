# ark_excel makefile
CC      ?= cc
CFLAGS  ?= -std=c99 -Wall -Wextra -I.
SRC      = src/ark_excel.c src/ark_zip.c src/ark_reader.c src/ark_inflate.c

all: demo

demo: examples/demo_report.c $(SRC)
	$(CC) $(CFLAGS) examples/demo_report.c $(SRC) -o demo

run: demo
	./demo && python tests/validate.py demo_report.xlsx

clean:
	rm -f demo demo_report.xlsx demo_report.exe

.PHONY: all run clean
