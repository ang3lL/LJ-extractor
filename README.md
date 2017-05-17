README
----

This dirty script extracts links embedded in
[Locky](https://www.hybrid-analysis.com/sample/cd0a031a65a10e8c549c29c1b5db87ad730c84ef9ba48041b3c4a723e56ee71f?environmentId=100)
or
[Jaff](https://www.hybrid-analysis.com/sample/83eb78af0b8697fcc6950a6f971245ba3ecf4a3182f5420b4f95d4e32840d0f1?environmentId=100)
PDFs using [PyMuPDF](https://github.com/rk700/PyMuPDF) and
[olevba](https://github.com/decalage2/oletools/wiki/olevba).

# Requirements

Install python requirements:

```
pip install -r requirements.txt
```

You also need to install [PyMuPDF](https://github.com/rk700/PyMuPDF#installation)

# Examples

Simple links extraction:

```
$ ./extract.py -v sample/nm.pdf 

trialinsider.com/f87346b
tiskr.com/f87346b
takanashi.jp/f87346b
boaevents.com/f87346b
```

Extraction with more verbosity:

```
$ ./extract.py -vv sample/001_5083.pdf 
[+] Embedded DLJTN2CUR.docm found
[+] Embedded file extracted: sample/001_5083.pdf-DLJTN2CUR.docm
[+] VBA Macros extracted
[+] Links found: ['dcfarbicka.sk/hHGFjd', 'hrlpk.com/hHGFjd', '5hdnnd74fffrottd.com/af/hHGFjd', 'byydei74fg43ff4f.net/af/hHGFjd', 'sjffonrvcik45bd.info/af/hHGFjd']
```

Massive extraction

```
$ for i in `ls sample/*.pdf`;do ./extract.py -v ${i};done | sort -u
5hdnnd74fffrottd.com/af/77g643
5hdnnd74fffrottd.com/af/f87346b
5hdnnd74fffrottd.com/af/hHGFjd
balprodukt.ru/77g643
beautyandearth.com/Nbiyure3
bellevillenorfolkterriers.co.uk/77g643
bianshop.com/hHGFjd
biarritzru.com/Nbiyure3
bioferme.biz/Nbiyure3
biolume.nl/77g643
bitsslab.com/77g643
bizcleaning.co.uk/hHGFjd
boaevents.com/f87346b
boolas.com/hHGFjd
byydei74fg43ff4f.net/af/77g643
byydei74fg43ff4f.net/af/f87346b
byydei74fg43ff4f.net/af/hHGFjd
daweizhi.com/Nbiyure3
dcfarbicka.sk/hHGFjd
demelkwegtuk.nl/77g643
diasgroup.sk/hHGFjd
djkammerthal.de/hHGFjd
dodawanie.com/Nbiyure3
...
```
