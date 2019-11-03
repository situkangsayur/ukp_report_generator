from xlsx import UkpManager


if __name__ == '__main__':
    t = UkpManager()
    t.load_data_xls('~/Documents/my wife/pkp/data.xls')
    print(t.ref_data)
    for temp in t.ref_data:
        print(str(temp))

    t.generate_using_template_xlsx('~/Documents/my wife/pkp/template2.xls', 'cp1252')

    #t.load_template_xlsx2('/home/hendri/Documents/my wife/pkp/jabar_pkp_sample.xlsx')

