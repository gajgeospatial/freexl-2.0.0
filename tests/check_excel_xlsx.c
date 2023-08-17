/* 
/ check_excel_xlsx.c
/
/ Test cases for Excel XLSX format files
/
/ version  2.0, 2021 June 06
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
	     "Sorry, this version of check_excel_xlsx was built by disabling support XML documents\n");
    return 0;
#else
    const void *handle;
    int ret;
    unsigned int info;
    const char *worksheet_name;
    unsigned short active_idx;
    unsigned int num_rows;
    unsigned short num_columns;
    FreeXL_CellValue cell_value;

    ret = freexl_open_xlsx ("testdata/test_xml.xlsx", &handle);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "OPEN ERROR: %d\n", ret);
	  return -1;
      }

    ret = freexl_get_worksheets_count (handle, &info);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr,
		   "GET_WORKSHEETS_COUNT ERROR for XLSX worksheet count: %d\n",
		   ret);
	  return -3;
      }
    if (info != 2)
      {
	  fprintf (stderr, "Unexpected XLSX worksheet count: %d\n", info);
	  return -4;
      }

    /* We start with the first worksheet, zero index */
    ret = freexl_get_worksheet_name (handle, 0, &worksheet_name);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting worksheet name: %d\n", ret);
	  return -5;
      }
    if (strcmp (worksheet_name, "First Sheet") != 0)
      {
	  fprintf (stderr, "Unexpected worksheet name: %s\n", worksheet_name);
	  return -6;
      }

    ret = freexl_select_active_worksheet (handle, 0);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error setting active worksheet: %d\n", ret);
	  return -7;
      }

    ret = freexl_get_active_worksheet (handle, &active_idx);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting active worksheet: %d\n", ret);
	  return -8;
      }
    if (active_idx != 0)
      {
	  fprintf (stderr, "Unexpected active sheet: %d\n", info);
	  return -9;
      }

    ret = freexl_worksheet_dimensions (handle, &num_rows, &num_columns);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting worksheet dimensions: %d\n", ret);
	  return -10;
      }
    if ((num_rows != 31) || (num_columns != 66))
      {
	  fprintf (stderr, "Unexpected active sheet dimensions: %u x %u\n",
		   num_rows, num_columns);
	  return -11;
      }

    ret = freexl_get_cell_value (handle, 0, 0, &cell_value);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting cell value (0,0): %d\n", ret);
	  return -12;
      }
    if (cell_value.type != FREEXL_CELL_SST_TEXT)
      {
	  fprintf (stderr, "Unexpected cell (0,0) type: %u\n", cell_value.type);
	  return -13;
      }
    if (strcmp (cell_value.value.text_value, "alpha") != 0)
      {
	  fprintf (stderr, "Unexpected cell (0,0) value: %s\n",
		   cell_value.value.text_value);
	  return -14;
      }

    ret = freexl_get_cell_value (handle, 1, 1, &cell_value);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting cell value (1,1): %d\n", ret);
	  return -15;
      }
    if (cell_value.type != FREEXL_CELL_INT)
      {
	  fprintf (stderr, "Unexpected cell (1,1) type: %u\n", cell_value.type);
	  return -16;
      }
    if (cell_value.value.int_value != 1)
      {
	  fprintf (stderr, "Unexpected cell (1,1) value: %d\n",
		   cell_value.value.int_value);
	  return -17;
      }

    ret = freexl_get_cell_value (handle, 2, 2, &cell_value);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting cell value (2,2): %d\n", ret);
	  return -18;
      }
    if (cell_value.type != FREEXL_CELL_DOUBLE)
      {
	  fprintf (stderr, "Unexpected cell (2,2) type: %u\n", cell_value.type);
	  return -19;
      }
    if (cell_value.value.double_value != 1.20)
      {
	  fprintf (stderr, "Unexpected cell (2,2) value: %f\n",
		   cell_value.value.double_value);
	  return -20;
      }

    ret = freexl_get_cell_value (handle, 3, 3, &cell_value);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting cell value (3,3): %d\n", ret);
	  return -21;
      }
    if (cell_value.type != FREEXL_CELL_DATE)
      {
	  fprintf (stderr, "Unexpected cell (3,3) type: %u\n", cell_value.type);
	  return -22;
      }
    if (strcmp (cell_value.value.text_value, "1956-02-02") != 0)
      {
	  fprintf (stderr, "Unexpected cell (3,3) value: %s\n",
		   cell_value.value.text_value);
	  return -23;
      }

    ret = freexl_get_cell_value (handle, 5, 1, &cell_value);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting cell value (5,1): %d\n", ret);
	  return -24;
      }
    if (cell_value.type != FREEXL_CELL_DOUBLE)
      {
	  fprintf (stderr, "Unexpected cell (5,1) type: %u\n", cell_value.type);
	  return -25;
      }
    if (cell_value.value.double_value != 1000000.50)
      {
	  fprintf (stderr, "Unexpected cell (5,1) value: %f\n",
		   cell_value.value.double_value);
	  return -26;
      }

    ret = freexl_get_cell_value (handle, 8, 1, &cell_value);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting cell value (8,1): %d\n", ret);
	  return -27;
      }
    if (cell_value.type != FREEXL_CELL_DOUBLE)
      {
	  fprintf (stderr, "Unexpected cell (8,1) type: %u\n", cell_value.type);
	  return -28;
      }
    if (cell_value.value.double_value != 1550.65)
      {
	  fprintf (stderr, "Unexpected cell (8,1) value: %f\n",
		   cell_value.value.double_value);
	  return -29;
      }

    ret = freexl_get_cell_value (handle, 9, 3, &cell_value);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting cell value (9,3): %d\n", ret);
	  return -30;
      }
    if (cell_value.type != FREEXL_CELL_DATETIME)
      {
	  fprintf (stderr, "Unexpected cell (9,3) type: %u\n", cell_value.type);
	  return -31;
      }
    if (strcmp (cell_value.value.text_value, "1956-02-02 19:23:00") != 0)
      {
	  fprintf (stderr, "Unexpected cell (9,3) value: %s\n",
		   cell_value.value.text_value);
	  return -32;
      }

    ret = freexl_get_cell_value (handle, 15, 1, &cell_value);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting cell value (15,1): %d\n", ret);
	  return -33;
      }
    if (cell_value.type != FREEXL_CELL_SST_TEXT)
      {
	  fprintf (stderr, "Unexpected cell (15,1) type: %u\n",
		   cell_value.type);
	  return -34;
      }
    if (strcmp (cell_value.value.text_value, "arezzo&firenze") != 0)
      {
	  fprintf (stderr, "Unexpected cell (15,1) value: %s\n",
		   cell_value.value.text_value);
	  return -35;
      }

    ret = freexl_get_cell_value (handle, 30, 65, &cell_value);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting cell value (30, 65): %d\n", ret);
	  return -36;
      }
    if (cell_value.type != FREEXL_CELL_SST_TEXT)
      {
	  fprintf (stderr, "Unexpected cell (30, 65) type: %u\n",
		   cell_value.type);
	  return -37;
      }
    if (strcmp (cell_value.value.text_value, "BN marker") != 0)
      {
	  fprintf (stderr, "Unexpected cell (30, 65) value: %s\n",
		   cell_value.value.text_value);
	  return -38;
      }

    /* testing the second worksheet, index 1 */
    ret = freexl_get_worksheet_name (handle, 1, &worksheet_name);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting worksheet name #2: %d\n", ret);
	  return -39;
      }
    if (strcmp (worksheet_name, "Second Sheet") != 0)
      {
	  fprintf (stderr, "Unexpected worksheet name #2: %s\n",
		   worksheet_name);
	  return -40;
      }

    ret = freexl_select_active_worksheet (handle, 1);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error setting active worksheet#2: %d\n", ret);
	  return -41;
      }

    ret = freexl_get_active_worksheet (handle, &active_idx);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting active worksheet#2: %d\n", ret);
	  return -42;
      }
    if (active_idx != 1)
      {
	  fprintf (stderr, "Unexpected active sheet #2: %d\n", info);
	  return -43;
      }

    ret = freexl_worksheet_dimensions (handle, &num_rows, &num_columns);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting worksheet dimensions #2: %d\n", ret);
	  return -44;
      }
    if ((num_rows != 14) || (num_columns != 8))
      {
	  fprintf (stderr, "Unexpected active sheet dimensions#2: %u x %u\n",
		   num_rows, num_columns);
	  return -45;
      }

    ret = freexl_get_cell_value (handle, 4, 2, &cell_value);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting cell value#2 (4,3): %d\n", ret);
	  return -46;
      }
    if (cell_value.type != FREEXL_CELL_SST_TEXT)
      {
	  fprintf (stderr, "Unexpected cell (4,2) type#2: %u\n",
		   cell_value.type);
	  return -47;
      }
    if (strcmp (cell_value.value.text_value, "some meaningless text") != 0)
      {
	  fprintf (stderr, "Unexpected cell (4,2) value: %s\n",
		   cell_value.value.text_value);
	  return -48;
      }

    /* error cases */
    ret = freexl_get_cell_value (handle, 1, 82, &cell_value);
    if (ret != FREEXL_ILLEGAL_CELL_ROW_COL)
      {
	  fprintf (stderr, "Unexpected result for (1,82): %d\n", ret);
	  return -49;
      }
    ret = freexl_get_cell_value (handle, 35, 6, &cell_value);
    if (ret != FREEXL_ILLEGAL_CELL_ROW_COL)
      {
	  fprintf (stderr, "Unexpected result for (35,6): %d\n", ret);
	  return -50;
      }
    ret = freexl_get_worksheet_name (handle, 10, &worksheet_name);
    if (ret == FREEXL_OK)
      {
	  fprintf (stderr, "Unexpected result getting worksheet name: %d\n",
		   ret);
	  return -51;
      }
    ret = freexl_select_active_worksheet (handle, 10);
    if (ret == FREEXL_OK)
      {
	  fprintf (stderr, "Unexpected result  setting active worksheet: %d\n",
		   ret);
	  return -52;
      }

/* testing GetStringsCount */
    ret = freexl_get_strings_count (handle, &info);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "Error getting XLSX Single Strings count: %d\n",
		   ret);
	  return -53;
      }
    if (info != 21)
      {
	  fprintf (stderr, "Unexpected XLSX worksheet count: %d\n", info);
	  return -54;
      }

    ret = freexl_close (handle);
    if (ret != FREEXL_OK)
      {
	  fprintf (stderr, "CLOSE ERROR: %d\n", ret);
	  return -2;
      }

    if (argc > 1 || argv[0] == NULL)
	argc = 1;		/* silencing stupid compiler warnings */

    return 0;
#endif /* end conditional XML support */
}
