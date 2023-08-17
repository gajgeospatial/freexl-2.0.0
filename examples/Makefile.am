noinst_PROGRAMS = test_xl test_xlsx test_ods xl2sql xlsx2sql ods2sql

AM_CFLAGS = -I@srcdir@/../headers
AM_LDFLAGS = -L../src -lfreexl -lm $(GCOV_FLAGS)

MOSTLYCLEANFILES = *.gcna *.gcno *.gcda

EXTRA_DIST = examples.doxy
