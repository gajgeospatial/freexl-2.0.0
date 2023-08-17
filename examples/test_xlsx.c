/* 
/ test_xlsx.c
/
/ FreeXL Sample code
/
/ Author: Sandro Furieri a.furieri@lqt.it
/
/ ------------------------------------------------------------------------------
/ 
/ Version: MPL 1.1/GPL 2.0/LGPL 2.1
/ 
/ The contents of this file are subject to the Mozilla Public License Version
/ 1.1 (the "License"); you may not use this file except in compliance with
/ the License. You may obtain a copy of the License at
/ http://www.mozilla.org/MPL/
/ 
/ Software distributed under the License is distributed on an "AS IS" basis,
/ WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
/ for the specific language governing rights and limitations under the
/ License.
/
/ The Original Code is the FreeXL library
/
/ The Initial Developer of the Original Code is Alessandro Furieri
/ 
/ Portions created by the Initial Developer are Copyright (C) 2021
/ the Initial Developer. All Rights Reserved.
/ 
/ Contributor(s):
/ 
/ Alternatively, the contents of this file may be used under the terms of
/ either the GNU General Public License Version 2 or later (the "GPL"), or
/ the GNU Lesser General Public License Version 2.1 or later (the "LGPL"),
/ in which case the provisions of the GPL or the LGPL are applicable instead
/ of those above. If you wish to allow use of your version of this file only
/ under the terms of either the GPL or the LGPL, and not to allow others to
/ use your version of this file under the terms of the MPL, indicate your
/ decision by deleting the provisions above and replace them with the notice
/ and other provisions required by the GPL or the LGPL. If you do not delete
/ the provisions above, a recipient may use your version of this file under
/ the terms of any one of the MPL, the GPL or the LGPL.
/ 
*/

#include <stdlib.h>
#include <stdio.h>
#include <string.h>

#include "freexl.h"

#if defined(_WIN32) && !defined(__MINGW32__)
#include "config-msvc.h"
#else
#include "config.h"
#endif

int
main (int argc, char *argv[])
{
#ifdef OMIT_XMLDOC		/* XML support is not enabled */
    fprintf (stderr,
	     "Sorry, this version of test_xlsx was built by disabling support XML documents\n");
    return -1;
#else
    const void *handle;
    int ret;
    unsigned int sst_count;
    unsigned int worksheet_count;
    unsigned int idx;
    const char *string;
    int verbose = 0;

    if (argc == 2 || argc == 3)
      {
	  if (argc == 3)
	    {
		if (strcmp (argv[2], "-verbose") == 0)
		    verbose = 1;
	    }
      }
    else
      {
	  fprintf (stderr, "usage: test_xlsx path.xlsx [-verbose]\n");
	  return -1;
      }

/* opening the .XLSX file [Workbook] */
    ret = freexl_open_xlsx (argv[1], &handle);
    if (ret != FREEXL_OK)
      {
	  freexl_close_xlsx (handle);
	  fprintf (stderr, "OPEN ERROR: %d\n", ret);
	  return -1;
      }

/* XLSX Worksheet entries */
    ret = freexl_get_worksheets_count (handle, &worksheet_count);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "GET-WORKSHEETS-COUNT Error: %d\n", ret);
	  goto stop;
      }
    printf ("XLSX Worksheets .....: %u\n", worksheet_count);

    /* XLSX SST entries */
    ret = freexl_get_strings_count (handle, &sst_count);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "GET-STRINGS-COUNT] Error: %d\n", ret);
	  goto stop;
      }
    printf ("XLSX SST entries ....: %u\n", sst_count);

    printf
	("\nWorksheets:\n=========================================================\n");
    for (idx = 0; idx < worksheet_count; idx++)
      {
	  /* printing XLSX Worksheets entries */
	  unsigned short active;
	  unsigned int rows;
	  unsigned short columns;

	  ret = freexl_get_worksheet_name (handle, idx, &string);
	  if (ret != FREEXL_OK)
	    {
		fprintf (stderr, "GET-WORKSHEET-NAME Error: %d\n", ret);
		goto stop;
	    }
	  if (string == NULL)
	      printf ("%3u] NULL (unnamed)\n", idx);
	  else
	      printf ("%3u] %s\n", idx, string);
	  ret = freexl_select_active_worksheet (handle, idx);
	  if (ret != FREEXL_OK)
	    {
		fprintf (stderr, "SELECT-ACTIVE_WORKSHEET Error: %d\n", ret);
		goto stop;
	    }
	  ret = freexl_get_active_worksheet (handle, &active);
	  if (ret != FREEXL_OK)
	    {
		fprintf (stderr, "GET-ACTIVE_WORKSHEET Error: %d\n", ret);
		goto stop;
	    }
	  printf
	      ("\tok, Worksheet successfully selected: currently active: %u\n",
	       active);
	  ret = freexl_worksheet_dimensions (handle, &rows, &columns);
	  if (ret != FREEXL_OK)
	    {
		fprintf (stderr, "WORKSHEET-DIMENSIONS Error: %d\n", ret);
		goto stop;
	    }
	  printf ("\t%u Rows X %u Columns\n\n", rows, columns);
      }

    if (!verbose)
	goto stop;

/* printing XLSX SST entries */
    printf
	("\nSST [Shared String Table]:\n=========================================================\n");
    for (idx = 0; idx < sst_count; idx++)
      {
	  ret = freexl_get_SST_string (handle, idx, &string);
	  if (ret != FREEXL_OK)
	    {
		fprintf (stderr, "GET-SST-STRING Error: %d\n", ret);
		goto stop;
	    }
	  if (string == NULL)
	      printf ("%8u] NULL (empty string)\n", idx);
	  else
	      printf ("%8u] %s\n", idx, string);
      }

  stop:
/* closing the .XLSX file [Workbook] */
    ret = freexl_close (handle);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "CLOSE ERROR: %d\n", ret);
	  return -1;
      }

    return 0;
#endif /* end conditional XML support */
}
