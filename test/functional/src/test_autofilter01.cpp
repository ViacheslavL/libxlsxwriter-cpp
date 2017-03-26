/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Test to compare output against Excel files.
 *
 * Copyright 2014-2016, John McNamara, jmcnamara@cpan.org
 *
 */

#include "xlsxwriter.hpp"

int main() {

    xlsxwriter::workbook  *workbook  = new xlsxwriter::workbook("test_autofilter01.xlsx");
    xlsxwriter::worksheet *worksheet = workbook->add_worksheet();
    uint16_t i;

    struct row {
        char region[16];
        char item[16];
        int  volume;
        char month[16];
    };

    struct row data[] = {
        {"East",  "Apple",   9000, "July"      },
        {"East",  "Apple",   5000, "July"      },
        {"South", "Orange",  9000, "September" },
        {"North", "Apple",   2000, "November"  },
        {"West",  "Apple",   9000, "November"  },
        {"South", "Pear",    7000, "October"   },
        {"North", "Pear",    9000, "August"    },
        {"West",  "Orange",  1000, "December"  },
        {"West",  "Grape",   1000, "November"  },
        {"South", "Pear",   10000, "April"     },
        {"West",  "Grape",   6000, "January"   },
        {"South", "Orange",  3000, "May"       },
        {"North", "Apple",   3000, "December"  },
        {"South", "Apple",   7000, "February"  },
        {"West",  "Grape",   1000, "December"  },
        {"East",  "Grape",   8000, "February"  },
        {"South", "Grape",  10000, "June"      },
        {"West",  "Pear",    7000, "December"  },
        {"South", "Apple",   2000, "October"   },
        {"East",  "Grape",   7000, "December"  },
        {"North", "Grape",   6000, "April"     },
        {"East",  "Pear",    8000, "February"  },
        {"North", "Apple",   7000, "August"    },
        {"North", "Orange",  7000, "July"      },
        {"North", "Apple",   6000, "June"      },
        {"South", "Grape",   8000, "September" },
        {"West",  "Apple",   3000, "October"   },
        {"South", "Orange", 10000, "November"  },
        {"West",  "Grape",   4000, "July"      },
        {"North", "Orange",  5000, "August"    },
        {"East",  "Orange",  1000, "November"  },
        {"East",  "Orange",  4000, "October"   },
        {"North", "Grape",   5000, "August"    },
        {"East",  "Apple",   1000, "December"  },
        {"South", "Apple",   10000, "March"    },
        {"East",  "Grape",   7000, "October"   },
        {"West",  "Grape",   1000, "September" },
        {"East",  "Grape",  10000, "October"   },
        {"South", "Orange",  8000, "March"     },
        {"North", "Apple",   4000, "July"      },
        {"South", "Orange",  5000, "July"      },
        {"West",  "Apple",   4000, "June"      },
        {"East",  "Apple",   5000, "April"     },
        {"North", "Pear",    3000, "August"    },
        {"East",  "Grape",   9000, "November"  },
        {"North", "Orange",  8000, "October"   },
        {"East",  "Apple",  10000, "June"      },
        {"South", "Pear",    1000, "December"  },
        {"North", "Grape",   10000, "July"     },
        {"East",  "Grape",   6000, "February"  }
    }; 


    /* Write the column headers. */
    worksheet->write_string(0, 0, "Region");
    worksheet->write_string(0, 1, "Item");
    worksheet->write_string(0, 2, "Volume");
    worksheet->write_string(0, 3, "Month");


    /* Write the row data. */
    for (i = 0; i < sizeof(data)/sizeof(struct row); i++) {
        worksheet->write_string(i + 1, 0, data[i].region);
        worksheet->write_string(i + 1, 1, data[i].item);
        worksheet->write_number(i + 1, 2, data[i].volume);
        worksheet->write_string(i + 1, 3, data[i].month);
    }
    

    worksheet->autofilter(0, 0, 50, 3);


    int result = workbook->close(); return result;
}
