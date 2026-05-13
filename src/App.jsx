import { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

// ── Sabitler ──────────────────────────────────────────────────────
const IDARELER = ["Tümü","KGM","DSİ","MEB","Çevre ve Şehircilik","Sağlık Bakanlığı","Belediye"];
const KATEGORILER = ["Tümü","Beton İşleri","Demir İşleri","Kaba İnşaat","İnce İnşaat","Tesisat","Elektrik","Zemin İşleri","Nakliye","Diğer"];

const _PD = [
[100,"15.100.1001","1 ton her cins çimento ve kirecin taşıtlara yükleme, boşaltma ve istifi","ton",256.25,280.21,"Nakliye","1)İhzaratta; işyerindeki malzemenin tartılmasından elde edilen ton cinsinden miktarıdır.",[["10.100.1062","Düz işçi","Sa",1.0,205.0],["19.100.1029","Yükleyici (100 HP)","Sa",0.021,1510.89]]],
[101,"15.100.1003","1 m³ her nevi taşın taşıtlara yükleme boşaltma ve figüresi","m³",43.44,57.94,"Nakliye","1)İhzaratta; işyerindeki malzemenin m³ cinsinden ölçülen miktarıdır.",[["19.100.1029","Yükleyici (100 HP)","Sa",0.023,1510.89],["19.100.1113","Mobil vinç","Sa",0.1,1402.45]]],
[102,"15.100.1005","1 ton çelik borunun taşıtlara yükleme, boşaltma ve istifi","ton",350.61,454.14,"Nakliye","",[["19.100.1113","Mobil vinç","Sa",0.2,1402.45],["19.100.1113","Mobil vinç","Sa",0.3,1402.45]]],
[103,"15.105.1001","Kazı alanı içine rastlayan fundaların gerekli el aletleri kullanarak kesilmesi ve temizlenmesi","m²",6406.25,7005.31,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",25.0,205.0],["19.100.1019","Traktör Buldozer (185 HP)","Sa",0.1,3141.71],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[104,"15.105.1002","Kazı ve dolgu alanında makine ile temizleme ve sökme işi yapılması","m²",648.96,811.55,"Zemin İşleri","",[["19.100.1019","Traktör Buldozer (185 HP)","Sa",0.1,3141.71],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[105,"15.105.1101","El ile ağaç kesilmesi ve sökme işi, çapı 5-10 cm (10 cm dahil) beher ağaç için","adet",128.13,140.11,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[106,"15.105.1102","Ağaç kesilmesi ve sökülmesi","adet",256.25,280.21,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[107,"15.105.1103","El ile ağaç kesilmesi ve sökme işi, çapı 21-30 cm (30 cm dahil) beher ağaç için","adet",512.5,560.43,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",2.0,205.0]]],
[108,"15.105.1104","El ile ağaç kesilmesi ve sökme işi, çapı 31-40 cm (40 cm dahil) beher ağaç için","adet",768.75,840.64,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",3.0,205.0]]],
[109,"15.105.1105","El ile ağaç kesilmesi ve sökme işi, çapı 41-50 cm (50 cm dahil) beher ağaç için","adet",1025.0,1120.85,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",4.0,205.0]]],
[110,"15.105.1106","El ile ağaç kesilmesi ve sökme işi, çapı 51-60 cm (60 cm dahil) beher ağaç için","adet",1537.5,1681.28,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",6.0,205.0]]],
[111,"15.105.1107","El ile ağaç kesilmesi ve sökme işi, çapı 61-70 cm (70 cm dahil) beher ağaç için","adet",2306.25,2521.91,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",9.0,205.0]]],
[112,"15.105.1108","El ile ağaç kesilmesi ve sökme işi, çapı 71-80 cm (80 cm dahil) beher ağaç için","adet",3075.0,3362.55,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",12.0,205.0]]],
[113,"15.105.1109","El ile ağaç kesilmesi ve sökme işi, çapı 81 cm den büyük olan beher ağaç için","adet",5125.0,5604.25,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",20.0,205.0]]],
[114,"15.106.1001","Patlayıcı Madde Kullanmadan Harçsız Kagir İnşaatın Yıkılması","m³",212.71,249.96,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: 1-Yıkım ya da sökümden çıkan işe yarar ve hurda değeri olan malzemenin sökümü biri",[["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.014,2183.46],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.022,1686.26],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[115,"15.106.1002","Patlayıcı Madde Kullanmadan Kireç ve Melez Harçlı Kagir İnşaatın Yıkılması","m³",270.03,324.14,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: 1-Yıkım ya da sökümden çıkan işe yarar ve hurda değeri olan malzemenin sökümü biri",[["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.035,2183.46],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.022,1686.26],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[116,"15.106.1003","Patlayıcı Madde Kullanmadan Çimento Harçlı Kargir ve Horasan İnşaatın Yıkılması","m³",365.55,447.75,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: 1-Yıkım ya da sökümden çıkan işe yarar ve hurda değeri olan malzemenin sökümü biri",[["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.07,2183.46],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.022,1686.26],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[117,"15.106.1004","Patlayıcı Madde Kullanmadan Demirli ve Demirsiz Beton İnşaatın Yıkılması","m³",671.24,843.34,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: 1-Yıkım ya da sökümden çıkan işe yarar ve hurda değeri olan malzemenin sökümü biri",[["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.182,2183.46],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.022,1686.26],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[118,"15.106.1005","İnsan Gücü ile (Kompresör vasıtasıyla) Her Cins Harçlı Kargir İnşaatın Yıkılması","m³",996.43,1129.71,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: 1-Yıkım ya da sökümden çıkan işe yarar ve hurda değeri olan malzemenin sökümü biri",[["10.100.1062","Düz işçi","Sa",0.75,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["19.100.1023","Kompresör","Sa",0.16,2508.08],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.022,1686.26]]],
[119,"15.106.1006","İnsan Gücü ile (Kompresör Vasıtasıyla) Demirli ve Demirsiz Beton İnşaatın Yıkılması","m³",1910.55,2159.71,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: 1-Yıkım ya da sökümden çıkan işe yarar ve hurda değeri olan malzemenin sökümü biri",[["10.100.1063","Erbab işçi","Sa",0.5,250.0],["10.100.1062","Düz işçi","Sa",1.5,205.0],["10.100.1062","Düz işçi","Sa",0.75,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["19.100.1023","Kompresör","Sa",0.32,2508.08],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.022,1686.26]]],
[120,"15.106.1007","Patlayıcı Madde Kullanarak Her Cins Harçlı Kargir ve Horasan İnşaatın Yıkılması","m³",1572.48,1719.13,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: 1-Yıkım ya da sökümden çıkan işe yarar ve hurda değeri olan malzemenin sökümü biri",[["10.160.1004","Fitil","m",0.5,12.0],["10.160.1005","Kapsül","adet",0.5,19.75],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",2.0,205.0],["10.100.1062","Düz işçi","Sa",2.5,205.0]]],
[121,"15.106.1008","Patlayıcı Madde Kullanarak Demirli ve Demirsiz Beton İnşaatın Yıkılması","m³",2367.19,2587.85,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: 1-Yıkım ya da sökümden çıkan işe yarar ve hurda değeri olan malzemenin sökümü biri",[["10.160.1004","Fitil","m",1.0,12.0],["10.160.1005","Kapsül","adet",1.0,19.75],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",2.0,310.0],["10.100.1062","Düz işçi","Sa",3.5,205.0],["10.100.1062","Düz işçi","Sa",2.5,205.0]]],
[122,"15.106.1011","Her Türlü Yıkımdan Çıkan Taş ve Tuğlanın İnsan Gücü ile Ayrılması","m³",768.75,840.64,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: 1- Yıkımdan çıkan malzemenin nihai ortalama 100 m'den fazla mesafeye taşınması dur",[["10.100.1062","Düz işçi","Sa",3.0,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[123,"15.106.1101","Doğal parke taşı, beton plak, adi kaldırım ve blokaj sökülmesi","m²",256.25,280.23,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: Sökümden çıkan malzemenin nihai ortalama 100 m'den fazla mesafeye taşınması durumu",[["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1062","Düz işçi","Sa",3.0,205.0],["10.100.1062","Düz işçi","Sa",2.5,205.0]]],
[124,"15.106.1102","Kırma taş, şose ve asfalt sökülmesi","m³",1409.38,1541.18,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: Sökümden çıkan malzemenin nihai ortalama 100 m'den fazla mesafeye taşınması durumu",[["10.100.1062","Düz işçi","Sa",3.0,205.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[125,"15.106.1103","Her türlü bordür sökülmesi","m",64.06,70.05,"Zemin İşleri","Birim Fiyata Dahil Olmayan Masraflar: Sökümden çıkan malzemenin nihai ortalama 100 m'den fazla mesafeye taşınması durumu",[["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[126,"15.106.1104","Her türlü iç sıva sökülmesi","m²",153.75,153.75,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1062","Düz işçi","Sa",0.1,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[127,"15.106.1105","Her türlü dış sıva sökülmesi","m²",281.88,281.88,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.1,205.0],["10.100.1062","Düz işçi","Sa",0.75,205.0],["10.100.1017","Dülger ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[128,"15.106.1106","Her türlü ahşap çatı sökülmesi","m²",514.06,514.06,"Zemin İşleri","ÖLÇÜ : Sökülen yatay alan üzerinden hesaplanır.",[["10.100.1062","Düz işçi","Sa",0.75,205.0],["10.100.1017","Dülger ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.27,205.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[129,"15.106.1107","Her türlü kiremit örtülü çatılarda kiremit aktarılması","m²",202.44,202.44,"Zemin İşleri","",[["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.27,205.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.1,310.0]]],
[130,"15.106.1108","Her türlü kiremit çatı örtüsü sökülmesi, toplanması, temizlenmesi, istif edilmesi","m²",115.63,115.63,"Zemin İşleri","",[["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.100.1026","Tenekeci ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[131,"15.106.1109","Galvanizli sac, alüminyum, cam elyaf takviyeli vb. çatı örtüsü sökülmesi","m²",180.0,180.0,"Zemin İşleri","",[["10.100.1026","Tenekeci ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0],["10.100.1026","Tenekeci ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[132,"15.106.1110","Kenetli sac ve bakır çatı örtüsü sökülmesi","m²",270.0,270.0,"Zemin İşleri","",[["10.100.1026","Tenekeci ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.100.1041","Marangoz usta yardımcısı","Sa",0.25,230.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[133,"15.106.1111","Çatı örtüsü altındaki ahşap kaplama tahtası sökülmesi","m²",200.0,200.0,"Zemin İşleri","",[["10.100.1041","Marangoz usta yardımcısı","Sa",0.25,230.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.1,300.0],["10.100.1062","Düz işçi","Sa",0.05,205.0]]],
[134,"15.106.1112","Çatı örtüleri altındaki su yalıtım örtüleri ve bitümlü karton sökülmesi","m²",50.31,50.31,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",0.1,300.0],["10.100.1062","Düz işçi","Sa",0.05,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.15,300.0],["10.100.1062","Düz işçi","Sa",0.15,205.0]]],
[135,"15.106.1113","Her türlü, sac, pvc, çinko vb, yağmur oluğu ve borusu sökülmesi","m",94.69,94.69,"Zemin İşleri","ÖLÇÜ : Sökülen malzemelerin uzunluğu yerinde ölçülerek hesaplanır.",[["10.100.1070","İkinci sınıf usta","Sa",0.15,300.0],["10.100.1062","Düz işçi","Sa",0.15,205.0],["10.100.1026","Tenekeci ustası","Sa",0.25,310.0],["10.100.1062","Düz işçi","Sa",0.35,205.0]]],
[136,"15.106.1115","Çinko ve sacdan her türlü atika duvar arkası çatı deresi (gizli dere) sökülmesi","m",308.75,308.75,"Zemin İşleri","",[["10.100.1026","Tenekeci ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.3,300.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[137,"15.106.1116","Her kalınlıkta mermer, traverten, terraza karo ve andezit kaplama sökülmesi","m²",266.25,266.25,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",0.3,300.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.4,300.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[138,"15.106.1117","Seramik, fayans vb. kaplama sökülmesi","m²",303.75,303.75,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",0.4,300.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.100.1070","İkinci sınıf usta","Sa",1.0,300.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[139,"15.106.1118","Dökme mozayik sökülmesi","m²",631.25,631.25,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",1.0,300.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.17,300.0],["10.100.1062","Düz işçi","Sa",0.34,205.0]]],
[140,"15.106.1119","PVC vb. döşeme kaplaması sökülmesi","m²",150.88,150.88,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",0.17,300.0],["10.100.1062","Düz işçi","Sa",0.34,205.0],["10.100.1041","Marangoz usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.8,205.0]]],
[141,"15.106.1120","Ahşap parke ve ahşap döşeme kaplaması sökülmesi","m²",320.0,320.0,"Zemin İşleri","",[["10.100.1041","Marangoz usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.8,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.07,300.0],["10.100.1062","Düz işçi","Sa",0.14,205.0],["10.100.1041","Marangoz usta yardımcısı","Sa",0.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[142,"15.106.1121","Ahşap ve PVC vb. süpürgelik sökülmesi","m",62.13,62.13,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",0.07,300.0],["10.100.1062","Düz işçi","Sa",0.14,205.0],["10.100.1041","Marangoz usta yardımcısı","Sa",0.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[143,"15.106.1122","Ahşap lambiri vb. duvar kaplaması sökülmesi","m²",400.0,400.0,"Zemin İşleri","",[["10.100.1041","Marangoz usta yardımcısı","Sa",0.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.5,300.0],["10.100.1062","Düz işçi","Sa",0.75,205.0]]],
[144,"15.106.1123","Mermer, traverten ve andezit plaktan denizlik, parapet ve harpuşta sökülmesi","m²",379.69,379.69,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",0.5,300.0],["10.100.1062","Düz işçi","Sa",0.75,205.0],["10.100.1070","İkinci sınıf usta","Sa",1.0,300.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.3,300.0]]],
[145,"15.106.1124","Dökme mozayikten denizlik, parapet ve harpuşta sökülmesi","m²",631.25,631.25,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",1.0,300.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.3,300.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[146,"15.106.1125","Duvar, döşeme ve tavanda sac, alüminyum, ahşap vb, dilatasyon fugası sökülmesi","m",266.25,266.25,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",0.3,300.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.100.1041","Marangoz usta yardımcısı","Sa",0.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[147,"15.106.1126","Her türlü ahşap dolap sökülmesi","m²",400.0,400.0,"Zemin İşleri","",[["10.100.1041","Marangoz usta yardımcısı","Sa",0.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.3,300.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.100.1041","Marangoz usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[148,"15.106.1127","Tezgah üstü mermeri sökülmesi","m²",266.25,266.25,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",0.3,300.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.100.1041","Marangoz usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[149,"15.106.1128","Ahşap asma tavan sökülmesi","m²",428.75,428.75,"Zemin İşleri","",[["10.100.1041","Marangoz usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.35,300.0],["10.100.1062","Düz işçi","Sa",0.7,205.0]]],
[150,"15.106.1129","Alüminyum, sac, alçıpanel, taşyünü, camyünü vb. asma tavan sökülmesi","m²",310.63,310.63,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",0.35,300.0],["10.100.1062","Düz işçi","Sa",0.7,205.0],["10.100.1070","İkinci sınıf usta","Sa",0.25,300.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1008","Doğramacı ustası","Sa",0.5,310.0]]],
[151,"15.106.1130","Alüminyum ve PVC den yapılan her türlü kapı ve pencere doğramasının sökülmesi","m²",221.88,221.88,"Zemin İşleri","",[["10.100.1070","İkinci sınıf usta","Sa",0.25,300.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1008","Doğramacı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[152,"15.106.1131","Her türlü ahşap kapı kasası, kapı kanadı, pencere ve camekan sökülmesi","m²",257.81,257.81,"Zemin İşleri","",[["10.100.1008","Doğramacı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.100.1018","Sıcak demirci ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[153,"15.110.1001","Her cins zeminde el ile yapılan (geniş-dar)derin kazılara derinlik zammı (iksasız kazılarda)","m³",0,0,"Zemin İşleri","m derinlikten sonra yapılan kazılara bu birim fiyat uygulanır.",[]],
[154,"15.115.1003","parçalanıp el ile atılabilen 0,100 m3 e kadar büyüklükteki her cins blok taşlar, kazı güçlüğü","m³",832.81,910.69,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",2.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[155,"15.115.1004","marn ve kil, 0,100-0,400 m³ büyüklükte parçalanıp el ile atılabilen her cins blok taşlar ve","m³",960.94,1050.8,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[156,"15.115.1005","kalker, marnlı kalker, marn, şist, gre, gevşek konglomera, alçı taşı, volkanik tüfler (bazalt","m³",768.35,851.0,"Zemin İşleri","",[["10.160.1004","Fitil","m",1.0,12.0],["10.160.1005","Kapsül","adet",1.0,19.75],["19.100.1023","Kompresör","Sa",0.06,2508.08],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.2,310.0],["10.100.1063","Erbab işçi","Sa",0.1,250.0],["10.100.1062","Düz işçi","Sa",1.25,205.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[157,"15.115.1006","halinde sert gre, betonlaşmış konglomera, kesif kalker, mermer, ayrışmamış serpantin,","m³",936.91,1040.94,"Zemin İşleri","",[["10.160.1004","Fitil","m",1.0,12.0],["10.160.1005","Kapsül","adet",1.0,19.75],["19.100.1023","Kompresör","Sa",0.09,2508.08],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.3,310.0],["10.100.1063","Erbab işçi","Sa",0.2,250.0],["10.100.1062","Düz işçi","Sa",1.25,205.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[158,"15.115.1007","ve benzeri, bazalt, profir, kuvars, 0,400 m³ den büyük aynı cins blok taşlar ve benzeri","m³",1137.66,1269.91,"Zemin İşleri","",[["10.160.1004","Fitil","m",1.0,12.0],["10.160.1005","Kapsül","adet",1.0,19.75],["19.100.1023","Kompresör","Sa",0.14,2508.08],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.4,310.0],["10.100.1063","Erbab işçi","Sa",0.2,250.0],["10.100.1062","Düz işçi","Sa",1.25,205.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[159,"15.115.1008","kalker, marnlı kalker, marn, şist, gre, gevşek konglomera, alçı taşı, volkanik tüfler (bazalt","m³",1306.09,1306.09,"Zemin İşleri","",[["19.100.1023","Kompresör","Sa",0.2,2508.08],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.5,310.0],["10.100.1063","Erbab işçi","Sa",0.2,250.0],["10.100.1062","Düz işçi","Sa",1.25,205.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[160,"15.115.1009","konglomera, kesif kalker, mermer, ayrışmamış serpantin, andezit, trakit bazalt tüfleri ve","m³",1776.35,1776.35,"Zemin İşleri","",[["19.100.1023","Kompresör","Sa",0.35,2508.08],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.5,310.0],["10.100.1063","Erbab işçi","Sa",0.2,250.0],["10.100.1062","Düz işçi","Sa",1.25,205.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[161,"15.115.1010","granit ve benzeri, bazalt, profir, kuvars, 0,400 m³ den büyük aynı cins blok taşlar ve benzeri","m³",2679.5,2679.5,"Zemin İşleri","",[["19.100.1023","Kompresör","Sa",0.6,2508.08],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.5,310.0],["10.100.1063","Erbab işçi","Sa",0.5,250.0],["10.100.1062","Düz işçi","Sa",1.25,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[162,"15.115.1201","toprak, gevşek silt, kum, kil, siltli, kumlu ve gevşek kil, killi kum ve çakıl, kürekle atılabilen","m³",960.94,960.94,"Zemin İşleri","",[["10.100.1062","Düz işçi","Sa",1.75,205.0],["10.100.1062","Düz işçi","Sa",0.75,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[163,"15.115.1202","toprak, gevşek silt, kum, kil, siltli, kumlu ve gevşek kil, killi kum ve çakıl, kürekle atılabilen","m³",1057.04,1057.04,"Zemin İşleri","",[["15.115.1201","","m³",1.1,768.75]]],
[164,"15.115.1211","derin kazı yapılması (ayrışmamış granit ve benzeri, bazalt, profir, kuvars, 0,400 m³ den","m³",3287.7,3287.7,"Zemin İşleri","",[["19.100.1023","Kompresör","Sa",0.7,2508.08],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.5,310.0],["10.100.1063","Erbab işçi","Sa",0.5,250.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",1.5,205.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[165,"15.115.1212","kazı yapılması (ayrışmamış granit ve benzeri, bazalt, profir, kuvars, 0,400 m³ den büyük","m³",3616.48,3616.48,"Zemin İşleri","",[["15.115.1211","kullanmadan her derinlikte çok sert kayada","m³",1.1,2630.16],["10.130.9991","Su","m³",0.1,55.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[166,"15.120.1001","Makine ile yumuşak ve sert toprak kazılması (Serbest kazı)","m³",62.35,83.45,"Zemin İşleri","",[["19.100.1006","Ekskavatör (paletli) (210-259 HP)","Sa",0.014,2971.46],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.005,1656.24]]],
[167,"15.120.1002","Makine ile yumuşak ve sert küskülük kazılması (serbest kazı)","m³",83.0,111.13,"Zemin İşleri","",[["19.100.1006","Ekskavatör (paletli) (210-259 HP)","Sa",0.019,2971.46],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.006,1656.24]]],
[168,"15.120.1003","Makine ile batak ve balçık kazılması (serbest kazı)","m³",122.64,164.18,"Zemin İşleri","",[["19.100.1006","Ekskavatör (paletli) (210-259 HP)","Sa",0.028,2971.46],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.009,1656.24]]],
[169,"15.120.1004","Makine ile patlayıcı madde kullanılarak yumuşak kaya kazılması (Serbest kazı)","m³",187.94,238.19,"Zemin İşleri","",[["19.100.1106","Paletli delgi makinası (160 HP)","Sa",0.007,3393.51],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.007,310.0],["10.100.1063","Erbab işçi","Sa",0.028,250.0],["19.100.1008","Ekskavatör (paletli) (260-299 HP)","Sa",0.022,3421.44],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.007,1656.24]]],
[170,"15.120.1005","Makine ile patlayıcı madde kullanmadan yumuşak kaya kazılması (Serbest kazı)","m³",138.39,187.91,"Zemin İşleri","",[["19.100.1008","Ekskavatör (paletli) (260-299 HP)","Sa",0.028,3421.44],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.009,1656.24]]],
[171,"15.120.1006","Makine ile patlayıcı madde kullanılarak sert kaya kazılması (Serbest kazı)","m³",250.45,250.45,"Zemin İşleri","",[["19.100.1106","Paletli delgi makinası (160 HP)","Sa",0.011,3393.51],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.011,310.0],["10.100.1063","Erbab işçi","Sa",0.044,250.0],["19.100.1008","Ekskavatör (paletli) (260-299 HP)","Sa",0.031,3421.44],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.01,1656.24]]],
[172,"15.120.1007","Makine ile patlayıcı madde kullanmadan sert kaya kazılması (Serbest kazı)","m³",318.26,318.26,"Zemin İşleri","",[["19.100.1009","Ekskavatör (paletli) (300-329 HP)","Sa",0.056,3984.58],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.019,1656.24]]],
[173,"15.120.1008","Makine ile patlayıcı madde kullanılarak çok sert kaya kazılması (Serbest kazı)","m³",328.28,328.28,"Zemin İşleri","",[["19.100.1106","Paletli delgi makinası (160 HP)","Sa",0.014,3393.51],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.014,310.0],["10.100.1063","Erbab işçi","Sa",0.056,250.0],["19.100.1008","Ekskavatör (paletli) (260-299 HP)","Sa",0.044,3421.44],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.015,1656.24]]],
[174,"15.120.1009","Makine ile patlayıcı madde kullanmadan çok sert kaya kazılması (Serbest kazı)","m³",430.3,430.3,"Zemin İşleri","",[["19.100.1009","Ekskavatör (paletli) (300-329 HP)","Sa",0.076,3984.58],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.025,1656.24]]],
[175,"15.120.1101","Makine ile her derinlik ve her genişlikte yumuşak ve sert toprak kazılması (Derin kazı)","m³",71.85,71.85,"Zemin İşleri","",[["19.100.1006","Ekskavatör (paletli) (210-259 HP)","Sa",0.016,2971.46],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.006,1656.24]]],
[176,"15.120.1102","Makine ile her derinlik ve her genişlikte yumuşak ve sert küskülük kazılması (Derin kazı)","m³",105.71,105.71,"Zemin İşleri","",[["19.100.1006","Ekskavatör (paletli) (210-259 HP)","Sa",0.024,2971.46],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.008,1656.24]]],
[177,"15.120.1103","Makine ile her derinlik ve her genişlikte batak ve balçık kazılması (Derin kazı)","m³",167.64,167.64,"Zemin İşleri","",[["19.100.1006","Ekskavatör (paletli) (210-259 HP)","Sa",0.039,2971.46],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.011,1656.24]]],
[178,"15.125.1001","Kum temin edilerek, el ile serme, sulama ve sıkıştırma yapılması","m³",588.13,635.06,"Zemin İşleri","",[["10.130.1004","Kum (elenmesi gerekmeyen ince agrega)","m³",1.0,137.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.130.9991","Su","m³",0.1,55.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.100.1062","Düz işçi","Sa",0.7,205.0],["10.130.1001","Çakıl (elenmesi gerekmeyen iri agrega)","m³",1.0,135.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.100.1062","Düz işçi","Sa",0.3,205.0]]],
[179,"15.125.1002","Çakıl temin edilerek, el ile serme, sulama ve sıkıştırma yapılması","m³",585.63,632.45,"Zemin İşleri","",[["10.130.1001","Çakıl (elenmesi gerekmeyen iri agrega)","m³",1.0,135.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.130.9991","Su","m³",0.1,55.0],["10.100.1062","Düz işçi","Sa",0.7,205.0],["10.130.1004","Kum (elenmesi gerekmeyen ince agrega)","m³",1.0,137.0],["19.100.1013","Greyder (190 HP)","Sa",0.01,2622.44],["10.130.9991","Su","m³",0.1,55.0]]],
[180,"15.125.1003","Kum temin edilerek, makine ile serme, sulama ve sıkıştırma yapılması","m³",257.78,294.56,"Zemin İşleri","",[["10.130.1004","Kum (elenmesi gerekmeyen ince agrega)","m³",1.0,137.0],["19.100.1013","Greyder (190 HP)","Sa",0.01,2622.44],["10.130.9991","Su","m³",0.1,55.0],["19.100.1044","Arazöz","Sa",0.013,882.41],["19.100.1047","Titreşimli Silindir (35-58 HP)","Sa",0.017,1531.16],["10.130.1001","Çakıl (elenmesi gerekmeyen iri agrega)","m³",1.0,135.0],["19.100.1013","Greyder (190 HP)","Sa",0.01,2622.44],["10.130.9991","Su","m³",0.1,55.0]]],
[181,"15.125.1004","Çakıl temin edilerek, makine ile serme, sulama ve sıkıştırma yapılması","m³",255.28,291.95,"Zemin İşleri","",[["10.130.1001","Çakıl (elenmesi gerekmeyen iri agrega)","m³",1.0,135.0],["19.100.1013","Greyder (190 HP)","Sa",0.01,2622.44],["10.130.9991","Su","m³",0.1,55.0],["19.100.1044","Arazöz","Sa",0.013,882.41],["19.100.1047","Titreşimli Silindir (35-58 HP)","Sa",0.017,1531.16],["10.130.1005","","m³",1.0,340.0],["10.100.1062","Düz işçi","Sa",2.0,205.0]]],
[182,"15.125.1005","Kum temin edilerek, drenaj yapılması","m³",937.5,937.5,"Zemin İşleri","",[["10.130.1005","","m³",1.0,340.0],["10.100.1062","Düz işçi","Sa",2.0,205.0],["10.130.1002","","m³",1.0,340.0],["10.100.1062","Düz işçi","Sa",2.0,205.0]]],
[183,"15.125.1006","Çakıl temin edilerek, drenaj yapılması","m³",937.5,937.5,"Zemin İşleri","",[["10.130.1002","","m³",1.0,340.0],["10.100.1062","Düz işçi","Sa",2.0,205.0],["10.130.1008","32 mm'ye kadar kırmataş","m³",1.0,485.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.130.9991","Su","m³",0.1,55.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.100.1062","Düz işçi","Sa",0.7,205.0]]],
[184,"15.125.1007","32mm'ye kadar kırmataş temin edilerek, el ile serme, sulama ve sıkıştırma yapılması","m³",1023.13,1023.13,"Zemin İşleri","",[["10.130.1008","32 mm'ye kadar kırmataş","m³",1.0,485.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.130.9991","Su","m³",0.1,55.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.100.1062","Düz işçi","Sa",0.7,205.0],["10.130.1008","32 mm'ye kadar kırmataş","m³",1.0,485.0],["19.100.1013","Greyder (190 HP)","Sa",0.01,2622.44],["10.130.9991","Su","m³",0.1,55.0]]],
[185,"15.125.1008","32mm'ye kadar kırmataş temin edilerek, makine ile serme, sulama ve sıkıştırma yapılması","m³",692.78,692.78,"Zemin İşleri","",[["10.130.1008","32 mm'ye kadar kırmataş","m³",1.0,485.0],["19.100.1013","Greyder (190 HP)","Sa",0.01,2622.44],["10.130.9991","Su","m³",0.1,55.0],["19.100.1044","Arazöz","Sa",0.013,882.41],["19.100.1047","Titreşimli Silindir (35-58 HP)","Sa",0.017,1531.16],["10.130.1009","","m³",1.0,425.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.130.9991","Su","m³",0.1,55.0]]],
[186,"15.125.1009","63 mm'ye kadar kırmataş temin edilerek, el ile serme, sulama ve sıkıştırma yapılması","m³",948.13,948.13,"Zemin İşleri","",[["10.130.1009","","m³",1.0,425.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.130.9991","Su","m³",0.1,55.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.100.1062","Düz işçi","Sa",0.7,205.0],["10.130.1009","","m³",1.0,425.0]]],
[187,"15.125.1010","63mm'ye kadar kırmataş temin edilerek, makine ile serme, sulama ve sıkıştırma yapılması","m³",617.78,617.78,"Zemin İşleri","",[["10.130.1009","","m³",1.0,425.0],["19.100.1013","Greyder (190 HP)","Sa",0.01,2622.44],["10.130.9991","Su","m³",0.1,55.0],["19.100.1044","Arazöz","Sa",0.013,882.41],["19.100.1047","Titreşimli Silindir (35-58 HP)","Sa",0.017,1531.16],["10.420.1852","Hafif agrega","m³",1.0,5.5],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[188,"15.125.1011","Hafif agrega (elenmiş kömür curufu) ile dolgu yapılması","m³",160.63,160.63,"Zemin İşleri","",[["10.420.1852","Hafif agrega","m³",1.0,5.5],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.130.4503","Yuvarlak yapı odunu","m³",0.025,5040.0],["10.100.1017","Dülger ustası","Sa",1.5,310.0],["10.100.1062","Düz işçi","Sa",1.5,205.0]]],
[189,"15.130.1001","Kazılara tam ahşap kaplamalı iksa yapılması","m²",1135.63,1135.63,"Zemin İşleri","I- Kazılara tam ahşap kaplamalı iksa yapılması: *** II- Kazılara sık aralıklı iksa yapılması, tüm iksa yüzünün en az %70",[["10.130.4503","Yuvarlak yapı odunu","m³",0.025,5040.0],["10.100.1017","Dülger ustası","Sa",1.5,310.0],["10.100.1062","Düz işçi","Sa",1.5,205.0]]],
[190,"15.130.1002","Kazılara tam ahşap kaplamalı iksa yapılması","m²",1135.63,1239.19,"Zemin İşleri","",[["15.130.1001","Kazılara tam ahşap kaplamalı iksa yapılması","m²",1.0,908.5],["15.130.1001","Kazılara tam ahşap kaplamalı iksa yapılması","m²",0.7,908.5]]],
[191,"15.130.1003","Kazılara sık aralıklı ahşap iksa yapılması","m²",794.94,867.44,"Zemin İşleri","",[["15.130.1001","Kazılara tam ahşap kaplamalı iksa yapılması","m²",0.7,908.5],["15.130.1001","Kazılara tam ahşap kaplamalı iksa yapılması","m²",0.5,908.5]]],
[192,"15.130.1004","Kazılara aralıklı ahşap kaplamalı iksa yapılması","m²",567.81,619.6,"Zemin İşleri","",[["15.130.1001","Kazılara tam ahşap kaplamalı iksa yapılması","m²",0.5,908.5],["19.100.1107","Jet grouting ekipmanı ile delgi makinası","Sa",0.075,9662.04],["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",0.2,445.0]]],
[193,"15.145.1001","basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (0,00 - 36,00 m","m³",7507.43,7507.43,"Zemin İşleri","",[["19.100.1115","Grab makinesi","Sa",0.05,15094.73],["19.100.1117","Bentonit ünitesi ve Desander (Grab için)","Sa",0.05,17460.28],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.05,355.0],["10.100.1060","Formen","Sa",0.05,445.0],["10.100.1063","Erbab işçi","Sa",0.05,250.0],["10.100.1062","Düz işçi","Sa",0.05,205.0]]],
[194,"15.145.1002","basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (36,00 - 72,00 m","m³",9620.55,9620.55,"Zemin İşleri","",[["19.100.1115","Grab makinesi","Sa",0.1,15094.73],["19.100.1117","Bentonit ünitesi ve Desander (Grab için)","Sa",0.1,17460.28],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.1,355.0],["10.100.1060","Formen","Sa",0.1,445.0],["10.100.1063","Erbab işçi","Sa",0.1,250.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[195,"15.145.1003","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (0,00 - 36,00","m³",9906.7,9906.7,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.04,59196.71],["19.100.1118","","Sa",0.04,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.04,355.0],["10.100.1060","Formen","Sa",0.04,445.0],["10.100.1063","Erbab işçi","Sa",0.04,250.0],["10.100.1062","Düz işçi","Sa",0.04,205.0]]],
[196,"15.145.1004","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (36,00 -","m³",12004.03,12004.03,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.06,59196.71],["19.100.1118","","Sa",0.06,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.06,355.0],["10.100.1060","Formen","Sa",0.06,445.0],["10.100.1063","Erbab işçi","Sa",0.06,250.0],["10.100.1062","Düz işçi","Sa",0.06,205.0]]],
[197,"15.145.1005","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (72,00 m","m³",18296.01,18296.01,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.12,59196.71],["19.100.1118","","Sa",0.12,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.12,355.0],["10.100.1060","Formen","Sa",0.12,445.0],["10.100.1063","Erbab işçi","Sa",0.12,250.0],["10.100.1062","Düz işçi","Sa",0.12,205.0]]],
[198,"15.145.1006","basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (0,00 - 36,00 m","m³",9620.55,9620.55,"Zemin İşleri","",[["19.100.1115","Grab makinesi","Sa",0.1,15094.73],["19.100.1117","Bentonit ünitesi ve Desander (Grab için)","Sa",0.1,17460.28],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.1,355.0],["10.100.1060","Formen","Sa",0.1,445.0],["10.100.1063","Erbab işçi","Sa",0.1,250.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[199,"15.145.1007","basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (36,00 - 72,00 m","m³",13001.55,13001.55,"Zemin İşleri","",[["19.100.1115","Grab makinesi","Sa",0.18,15094.73],["19.100.1117","Bentonit ünitesi ve Desander (Grab için)","Sa",0.18,17460.28],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.18,355.0],["10.100.1060","Formen","Sa",0.18,445.0],["10.100.1063","Erbab işçi","Sa",0.18,250.0],["10.100.1062","Düz işçi","Sa",0.18,205.0]]],
[200,"15.145.1008","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (0,00 - 36,00","m³",12004.03,12004.03,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.06,59196.71],["19.100.1118","","Sa",0.06,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.06,355.0],["10.100.1060","Formen","Sa",0.06,445.0],["10.100.1063","Erbab işçi","Sa",0.06,250.0],["10.100.1062","Düz işçi","Sa",0.06,205.0]]],
[201,"15.145.1009","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (36,00 -","m³",16198.69,16198.69,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.1,59196.71],["19.100.1118","","Sa",0.1,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.1,355.0],["10.100.1060","Formen","Sa",0.1,445.0],["10.100.1063","Erbab işçi","Sa",0.1,250.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[202,"15.145.1010","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (72,00 m","m³",23539.33,23539.33,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.17,59196.71],["19.100.1118","","Sa",0.17,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.17,355.0],["10.100.1060","Formen","Sa",0.17,445.0],["10.100.1063","Erbab işçi","Sa",0.17,250.0],["10.100.1062","Düz işçi","Sa",0.17,205.0]]],
[203,"15.145.1011","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (0,00 - 36,00","m³",16198.69,16198.69,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.1,59196.71],["19.100.1118","","Sa",0.1,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.1,355.0],["10.100.1060","Formen","Sa",0.1,445.0],["10.100.1063","Erbab işçi","Sa",0.1,250.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[204,"15.145.1012","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (36,00 -","m³",23539.33,23539.33,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.17,59196.71],["19.100.1118","","Sa",0.17,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.17,355.0],["10.100.1060","Formen","Sa",0.17,445.0],["10.100.1063","Erbab işçi","Sa",0.17,250.0],["10.100.1062","Düz işçi","Sa",0.17,205.0]]],
[205,"15.145.1013","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (72,00 m","m³",32977.3,32977.3,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.26,59196.71],["19.100.1118","","Sa",0.26,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.26,355.0],["10.100.1060","Formen","Sa",0.26,445.0],["10.100.1063","Erbab işçi","Sa",0.26,250.0],["10.100.1062","Düz işçi","Sa",0.26,205.0]]],
[206,"15.145.1014","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (0,00 - 36,00","m³",24588.0,24588.0,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.18,59196.71],["19.100.1118","","Sa",0.18,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.18,355.0],["10.100.1060","Formen","Sa",0.18,445.0],["10.100.1063","Erbab işçi","Sa",0.18,250.0],["10.100.1062","Düz işçi","Sa",0.18,205.0]]],
[207,"15.145.1015","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (36,00 -","m³",32977.3,32977.3,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.26,59196.71],["19.100.1118","","Sa",0.26,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.26,355.0],["10.100.1060","Formen","Sa",0.26,445.0],["10.100.1063","Erbab işçi","Sa",0.26,250.0],["10.100.1062","Düz işçi","Sa",0.26,205.0]]],
[208,"15.145.1016","C30/35 basınç dayanımımda yerinde dökme beton ile diyafram duvar yapılması (72,00 m","m³",45561.26,45561.26,"Zemin İşleri","",[["19.100.1116","Hidrofreze makinesi","Sa",0.38,59196.71],["19.100.1118","","Sa",0.38,23441.37],["10.450.9601","Bentonit","ton",0.0275,4500.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1055","Operatör makinist","Sa",0.38,355.0],["10.100.1060","Formen","Sa",0.38,445.0],["10.100.1063","Erbab işçi","Sa",0.38,250.0],["10.100.1062","Düz işçi","Sa",0.38,205.0]]],
[209,"15.150.5101","Kayar Kalıplı Beton Finişeri ile Yol Betonu Serilmesi (C30/37 Beton Dayanım Sınıfında)","m³",5357.55,5357.55,"Beton İşleri","",[["10.130.1506","C 30/37 beton harcı","m³",1.0,2700.0],["19.100.3003","","Sa",0.0048,37189.8],["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.009,2183.46],["10.100.1015","Betoncu ustası","Sa",1.364,310.0],["10.100.1062","Düz işçi","Sa",1.364,205.0],["10.100.1060","Formen","Sa",0.341,445.0],["10.100.1061","Topoğraf","Sa",1.364,340.0],["10.100.1062","Düz işçi","Sa",0.341,205.0]]],
[210,"15.150.5102","Kayar Kalıplı Beton Finişeri ile Yol Betonu Serilmesi (C35/40 Beton Dayanım Sınıfında)","m³",5545.05,5545.05,"Beton İşleri","",[["10.130.1507","C 35/45 beton harcı","m³",1.0,2850.0],["19.100.3003","","Sa",0.0048,37189.8],["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.009,2183.46],["10.100.1015","Betoncu ustası","Sa",1.364,310.0],["10.100.1062","Düz işçi","Sa",1.364,205.0],["10.100.1060","Formen","Sa",0.341,445.0],["10.100.1061","Topoğraf","Sa",1.364,340.0],["10.100.1062","Düz işçi","Sa",0.341,205.0]]],
[211,"15.150.5103","Kayar Kalıplı Beton Finişeri ile Yol Betonu Serilmesi (C40/50 Beton Dayanım Sınıfında)","m³",5670.05,5670.05,"Beton İşleri","",[["10.130.1508","C 40/50 beton harcı","m³",1.0,2950.0],["19.100.3003","","Sa",0.0048,37189.8],["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.009,2183.46],["10.100.1015","Betoncu ustası","Sa",1.364,310.0],["10.100.1062","Düz işçi","Sa",1.364,205.0],["10.100.1060","Formen","Sa",0.341,445.0],["10.100.1061","Topoğraf","Sa",1.364,340.0],["10.100.1062","Düz işçi","Sa",0.341,205.0]]],
[212,"15.150.5111","Beton Yol Yüzeyinin Elle Kürlenmesi, Perdahlanması ve Pürüzlendirilmesi","m²",90.63,90.63,"Beton İşleri","",[["10.100.1015","Betoncu ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[213,"15.165.1001","kullanılan münferit çatı aşıkları ve mertekleri, lentolar, hurdi döşemeler, köşe takviye","ton",85414.38,93386.28,"Demir İşleri","1) Kullanılan profil demiri tespit malzemesiyle birlikte boyanmadan önce tartılır.",[["19.100.1089","Demir doğrama imalat atelyesi","Sa",7.0,5235.8],["19.100.1113","Mobil vinç","Sa",2.0,1402.45],["10.100.1019","Soğuk demirci ustası","Sa",2.0,310.0],["10.100.1062","Düz işçi","Sa",2.0,205.0]]],
[214,"15.165.1002","Profil demirlerinden çatı makası yapılması ve yerine konulması.","ton",92835.66,101693.96,"Demir İşleri","",[["19.100.1089","Demir doğrama imalat atelyesi","Sa",8.0,5235.8],["19.100.1113","Mobil vinç","Sa",2.5,1402.45],["10.100.1018","Sıcak demirci ustası","Sa",2.0,310.0],["10.100.1062","Düz işçi","Sa",2.0,205.0]]],
[215,"15.165.1003","(yapı karkası, köprülerde profil demirlerinden kirişler, başlıklar, bağlantılar ve benzeri","ton",85943.51,85943.51,"Demir İşleri","1) Ölçmede tartı esastır, kullanılan profil demiri perçin, cıvata, ek levhaları ve benzeri tespit elemanları birlikte bo",[["19.100.1089","Demir doğrama imalat atelyesi","Sa",7.0,5235.8],["19.100.1113","Mobil vinç","Sa",1.8,1402.45],["10.100.1019","Soğuk demirci ustası","Sa",1.8,310.0],["10.100.1062","Düz işçi","Sa",1.8,205.0]]],
[216,"15.180.1001","Ahşaptan seri kalıp yapılması","m²",331.63,362.23,"Beton İşleri","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.005,9000.0],["19.100.1091","","Sa",0.005,6059.7],["10.100.1086","Ahşap Kalıpçı (Betonarme)","Sa",0.3,310.0],["10.100.1090","Kalıp İşleri Usta Yardımcısı","Sa",0.3,230.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[217,"15.180.1002","Ahşaptan düz yüzeyli beton ve betonarme kalıbı yapılması","m²",873.03,953.18,"Beton İşleri","kalıpları ölçüye dâhil edilmez. Deliğin kalıp tarafındaki yüzünden delik boşluğu çıkarılmaz.",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.0123,9000.0],["19.100.1091","","Sa",0.012,6059.7],["10.100.1086","Ahşap Kalıpçı (Betonarme)","Sa",0.75,310.0],["10.100.1090","Kalıp İşleri Usta Yardımcısı","Sa",0.75,230.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[218,"15.180.1003","Plywood ile düz yüzeyli betonarme kalıbı yapılması","m²",1028.33,1123.51,"Beton İşleri","kalıpları ölçüye dâhil edilmez. Deliğin kalıp tarafındaki yüzünden delik boşluğu çıkarılmaz.",[["10.130.4606","","m²",0.04,609.0],["10.130.4606","","m²",0.004,609.0],["10.130.4607","I kesitli ahşap kiriş","m",0.11,301.0],["10.100.1088","Panel Kalıpçı (Betonarme)","Sa",1.1,310.0],["10.100.1090","Kalıp İşleri Usta Yardımcısı","Sa",1.1,230.0],["10.100.1062","Düz işçi","Sa",0.75,205.0]]],
[219,"15.180.1004","Sac ile eğri yüzeyli beton ve betonarme kalıbı yapılması","m²",983.15,983.15,"Beton İşleri","kalıpları ölçüye dâhil edilmez. Deliğin kalıp tarafındaki yüzünden delik boşluğu çıkarılmaz.",[["19.100.1089","Demir doğrama imalat atelyesi","Sa",0.025,5235.8],["19.100.1113","Mobil vinç","Sa",0.18,1402.45],["10.100.1089","Metal Kalıp Ustası (Betonarme)","Sa",0.5,310.0],["10.100.1090","Kalıp İşleri Usta Yardımcısı","Sa",0.5,230.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[220,"15.180.1007","Tünel kalıp sistemi ile betonarme kalıp yapılması","m²",0,0,"Beton İşleri","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.001,9000.0],["19.100.1090","Tünel kalıp imalat atelyesi","Sa",0.015,5281.3],["19.100.1105","Kule vinç","Sa",0.023,3590.32],["19.100.1104","Mobil vinç (60ton, 240 HP)","Sa",0.002,4850.26],["10.100.1068","Birinci sınıf usta","Sa",0.03,310.0],["10.100.1018","Sıcak demirci ustası","Sa",0.01,310.0],["10.100.1062","Düz işçi","Sa",0.01,205.0],["10.100.1060","Formen","Sa",0.01,445.0]]],
[221,"15.185.1005","Çelik borudan kalıp iskelesi yapılması (0,00-4,00 m arası)","m³",127.75,127.75,"Beton İşleri","1) Bu ölçü kapsamına giren yapı ve sınaî imalâtın kalıp gören yüzü ile iskelenin isnat ettiği zemin arasındaki boşluk he",[["19.100.1120","","Sa",0.004,3900.9],["10.100.1035","İskele Kurulum Elemanı","Sa",0.24,310.0]]],
[222,"15.185.1006","Çelik borudan kalıp iskelesi yapılması (4,01-6,00 m arası)","m³",147.68,147.68,"Beton İşleri","1) Bu ölçü kapsamına giren yapı ve sınaî imalâtın kalıp gören yüzü ile iskelenin isnat ettiği zemin arasındaki boşluk he",[["19.100.1120","","Sa",0.004,3900.9],["10.100.1035","İskele Kurulum Elemanı","Sa",0.28,310.0]]],
[223,"15.185.1007","Çelik borudan kalıp iskelesi yapılması (6,01-8,00m arası)","m³",167.61,167.61,"Beton İşleri","1) Bu ölçü kapsamına giren yapı ve sınaî imalâtın kalıp gören yüzü ile iskelenin isnat ettiği zemin arasındaki boşluk he",[["19.100.1120","","Sa",0.004,3900.9],["10.100.1035","İskele Kurulum Elemanı","Sa",0.32,310.0]]],
[224,"15.185.1008","Çelik borudan kalıp iskelesi yapılması 8,01-10,00m arası)","m³",187.55,187.55,"Beton İşleri","1) Bu ölçü kapsamına giren yapı ve sınaî imalâtın kalıp gören yüzü ile iskelenin isnat ettiği zemin arasındaki boşluk he",[["19.100.1120","","Sa",0.004,3900.9],["10.100.1035","İskele Kurulum Elemanı","Sa",0.36,310.0]]],
[225,"15.185.1041","Güvenlik Ağı Yapılması Sistem T (TS EN 1263-1)","m²",38.6,38.6,"Beton İşleri","",[["10.130.4721","Güvenlik Ağı Sistem T","m²",0.15,120.0],["10.100.1035","İskele Kurulum Elemanı","Sa",0.025,310.0],["10.100.1062","Düz işçi","Sa",0.025,205.0]]],
[226,"15.185.1042","Güvenlik Ağı Yapılması (Sistem S ve Sistem U) TS EN 1263-1","m²",38.6,38.6,"Beton İşleri","",[["10.130.4722","Güvenlik Ağı Sistem S ve Sistem U","m²",0.15,120.0],["10.100.1035","İskele Kurulum Elemanı","Sa",0.025,310.0],["10.100.1062","Düz işçi","Sa",0.025,205.0]]],
[227,"15.190.1001","Bazalt agregalı (gri) yüzey sertlestirici ve kür uygulaması (taze betonda)","m²",176.13,176.13,"Beton İşleri","",[["19.100.1094","Perdah Makinesi","Sa",0.1,483.01],["10.100.1015","Betoncu ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[228,"15.190.1002","Kuvars agregalı (gri) yüzey sertlestirici ve kür uygulaması (taze betonda)","m²",179.25,179.25,"Beton İşleri","",[["19.100.1094","Perdah Makinesi","Sa",0.1,483.01],["10.100.1015","Betoncu ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[229,"15.190.1003","Kuvars-Korund agregalı (gri) yüzey sertlestirici ve kür uygulaması (taze betonda)","m²",194.88,194.88,"Beton İşleri","",[["19.100.1094","Perdah Makinesi","Sa",0.1,483.01],["10.100.1015","Betoncu ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[230,"15.190.1004","Korund agregalı (gri) yüzey sertlestirici ve kür uygulaması (taze betonda)","m²",207.38,207.38,"Beton İşleri","",[["19.100.1094","Perdah Makinesi","Sa",0.1,483.01],["10.100.1015","Betoncu ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[231,"15.190.1006","Yeni beton yüzeylere kür yapılması (Saha Betonu)","m²",36.38,36.38,"Beton İşleri","",[["10.100.1015","Betoncu ustası","Sa",0.05,310.0]]],
[232,"15.190.1008","Silindir ile Sıkıştrılmış Beton Yollarda Parafinik Esaslı Kür Malzemesi ile Kür Yapılması","m²",57.25,57.25,"Beton İşleri","",[["10.100.1015","Betoncu ustası","Sa",0.08,310.0],["10.100.1015","Betoncu ustası","Sa",0.08,310.0]]],
[233,"15.190.1009","Silindir ile Sıkıştrılmış Beton Yollarda Akrilik Esaslı Kür Malzemesi ile Kür Yapılması","m²",60.75,60.75,"Beton İşleri","",[["10.100.1015","Betoncu ustası","Sa",0.08,310.0],["10.130.9991","Su","m³",3.0,55.0],["19.100.1044","Arazöz","Sa",0.39,882.41]]],
[234,"15.190.1012","2,5 mm kalınlıkta, self leveling Poliüretan esaslı zemin kaplaması yapılması","m²",2123.96,2123.96,"Beton İşleri","",[["19.100.1085","Karıştırıcı","Sa",0.075,7.59],["19.100.1032","Mozayik Silme Makinası","Sa",0.1,528.95],["10.100.1068","Birinci sınıf usta","Sa",0.5,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.5,230.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[235,"15.190.1014","2,5 mm kalınlıkta self leveling Epoksi esaslı zemin kaplaması yapılması","m²",1551.71,1551.71,"Beton İşleri","",[["19.100.1085","Karıştırıcı","Sa",0.075,7.59],["19.100.1032","Mozayik Silme Makinası","Sa",0.1,528.95],["10.100.1068","Birinci sınıf usta","Sa",0.5,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.5,230.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[236,"15.190.1015","bakteriostatik, iki bileşenli, poliüretan esaslı, şeffaf veya pigmentli, mat yüzey bitişli son kat","m²",231.25,231.25,"Beton İşleri","",[["10.100.1068","Birinci sınıf usta","Sa",0.1,310.0],["10.100.1068","Birinci sınıf usta","Sa",0.1,310.0]]],
[237,"15.190.1016","uygulaması üzerine, anti-statik, iki bileşenli, poliüretan esaslı, mat, su bazlı ve düşük","m²",460.63,460.63,"Beton İşleri","",[["10.100.1068","Birinci sınıf usta","Sa",0.1,310.0]]],
[238,"15.190.1019","Alçı esaslı kendiliğinden yerleşen harç ile ortalama 2 mm kalınlıkta zemin tesviyesi yapılması","m²",176.09,176.09,"Beton İşleri","",[["19.100.1085","Karıştırıcı","Sa",0.05,7.59],["10.130.9991","Su","m³",0.001,55.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[239,"15.200.1001","esaslı drenaj ve koruma levhası temini ve yerine döşenmesi","m²",110.0,110.0,"Kaba İnşaat","",[["10.330.6401","HDPE esaslı koruma ve drenaj levhası","m²",1.05,25.0],["10.100.1010","Yalıtımcı ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.15,205.0]]],
[240,"15.200.1002","esaslı drenaj ve koruma levhası temini ve yerine döşenmesi","m²",133.63,133.63,"Kaba İnşaat","",[["10.330.6402","HDPE esaslı koruma ve drenaj levhası","m²",1.05,43.0],["10.100.1010","Yalıtımcı ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.15,205.0]]],
[241,"15.200.1003","esaslı drenaj ve koruma levhası temini ve yerine döşenmesi","m²",158.56,158.56,"Kaba İnşaat","",[["10.330.6403","HDPE esaslı koruma ve drenaj levhası","m²",1.05,62.0],["10.100.1010","Yalıtımcı ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.15,205.0]]],
[242,"15.200.1004","yerine döşenmesi","m²",130.31,130.31,"Kaba İnşaat","",[["10.330.6401","HDPE esaslı koruma ve drenaj levhası","m²",1.05,25.0],["10.100.1010","Yalıtımcı ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[243,"15.200.1005","yerine döşenmesi","m²",153.94,153.94,"Kaba İnşaat","",[["10.330.6402","HDPE esaslı koruma ve drenaj levhası","m²",1.05,43.0],["10.100.1010","Yalıtımcı ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[244,"15.200.1006","yerine döşenmesi","m²",178.88,178.88,"Kaba İnşaat","",[["10.330.6403","HDPE esaslı koruma ve drenaj levhası","m²",1.05,62.0],["10.100.1010","Yalıtımcı ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.440.1022","Ø 100 mm anma çaplı","m",1.0,32.0]]],
[245,"15.204.1001","Ø 100 mm anma çaplı, PVC esaslı koruge drenaj borusunun temini ve yerine döşenmesi","m",56.0,56.0,"Tesisat","",[["10.440.1022","Ø 100 mm anma çaplı","m",1.0,32.0],["10.100.1049","Borucu usta yardımcısı","Sa",0.02,230.0],["10.100.1062","Düz işçi","Sa",0.04,205.0],["10.440.1023","Ø 125 mm anma çaplı","m",1.0,55.0],["10.100.1049","Borucu usta yardımcısı","Sa",0.02,230.0]]],
[246,"15.204.1002","Ø 125 mm anma çaplı, PVC esaslı koruge drenaj borusunun temini ve yerine döşenmesi","m",84.75,84.75,"Tesisat","",[["10.440.1023","Ø 125 mm anma çaplı","m",1.0,55.0],["10.100.1049","Borucu usta yardımcısı","Sa",0.02,230.0],["10.100.1062","Düz işçi","Sa",0.04,205.0],["10.440.1024","Ø 160 mm anma çaplı","m",1.0,85.0],["10.100.1049","Borucu usta yardımcısı","Sa",0.02,230.0],["10.100.1062","Düz işçi","Sa",0.04,205.0]]],
[247,"15.204.1003","Ø 160 mm anma çaplı, PVC esaslı koruge drenaj borusunun temini ve yerine döşenmesi","m",122.25,122.25,"Tesisat","",[["10.440.1024","Ø 160 mm anma çaplı","m",1.0,85.0],["10.100.1049","Borucu usta yardımcısı","Sa",0.02,230.0],["10.100.1062","Düz işçi","Sa",0.04,205.0],["10.440.1025","Ø 200 mm anma çaplı","m",1.0,120.0],["10.100.1049","Borucu usta yardımcısı","Sa",0.02,230.0]]],
[248,"15.204.1004","Ø 200 mm anma çaplı, PVC esaslı koruge drenaj borusunun temini ve yerine döşenmesi","m",166.0,166.0,"Tesisat","",[["10.440.1025","Ø 200 mm anma çaplı","m",1.0,120.0],["10.100.1049","Borucu usta yardımcısı","Sa",0.02,230.0],["10.100.1062","Düz işçi","Sa",0.04,205.0],["19.100.2015","Ocakta hazırlanan taş","m³",1.25,345.39],["10.100.1013","Duvarcı ustası","Sa",0.75,310.0],["10.100.1045","Duvarcı usta yardımcısı","Sa",1.0,230.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[249,"15.210.1001","Ocak taşı ile kuru duvar inşaat yapılması","m³",1496.54,1496.54,"Kaba İnşaat","",[["19.100.2015","Ocakta hazırlanan taş","m³",1.25,345.39],["10.100.1013","Duvarcı ustası","Sa",0.75,310.0],["10.100.1045","Duvarcı usta yardımcısı","Sa",1.0,230.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.1,1656.24]]],
[250,"15.210.1004","Ocak taşı ile blokaj yapılması","m³",1412.41,1412.41,"Kaba İnşaat","",[["19.100.2015","Ocakta hazırlanan taş","m³",1.1,345.39],["10.100.1013","Duvarcı ustası","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",1.5,205.0],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.08,1656.24],["19.100.2015","Ocakta hazırlanan taş","m³",1.2,345.39],["19.100.2503","200 Kg çimento dozlu harç yapılması","m³",0.33,1578.5]]],
[251,"15.210.1005","Ocak taşı ile 200 dozlu çimento harçlı kargir inşaat yapılması","m³",2384.78,2384.78,"Kaba İnşaat","1) Projesindeki boyutlar üzerinden hesaplanır.",[["19.100.2015","Ocakta hazırlanan taş","m³",1.2,345.39],["19.100.2503","200 Kg çimento dozlu harç yapılması","m³",0.33,1578.5],["10.100.1013","Duvarcı ustası","Sa",1.0,310.0],["10.100.1045","Duvarcı usta yardımcısı","Sa",1.25,230.0],["10.100.1062","Düz işçi","Sa",0.75,205.0],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.1,1656.24]]],
[252,"15.210.1006","Ocaktan çaplanmış moloz taşı ile 200 dozlu çimento harçlı kargir inşaat yapılması","m³",3911.66,3911.66,"Kaba İnşaat","1) Projesindeki boyutlar üzerinden hesaplanır.",[["19.100.2020","Ocakta çaplanmış moloz taş","m³",1.1,1400.39],["19.100.2503","200 Kg çimento dozlu harç yapılması","m³",0.25,1578.5],["10.100.1013","Duvarcı ustası","Sa",1.25,310.0],["10.100.1045","Duvarcı usta yardımcısı","Sa",1.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["19.100.1027","Kazıcı yükleyici (100 HP)","Sa",0.1,1656.24]]],
[253,"15.220.1011","85 mm kalınlığında yatay delikli tuğla (190 x 85 x 190 mm) ile duvar yapılması","m²",730.71,792.25,"Kaba İnşaat","",[["10.130.2001","190 x 85 x 190 mm yatay delikli tuğla","adet",26.0,5.1],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.01,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",1.2,205.0],["10.130.2004","200 x 100 x 200 mm yatay delikli tuğla","adet",24.0,5.9],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.011,1941.51],["10.130.9991","Su","m³",0.01,55.0]]],
[254,"15.220.1012","100 mm kalınlığında yatay delikli tuğla (200 x 100 x 200 mm) ile duvar yapılması","m²",753.39,816.7,"Kaba İnşaat","",[["10.130.2004","200 x 100 x 200 mm yatay delikli tuğla","adet",24.0,5.9],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.011,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.61,310.0],["10.100.1062","Düz işçi","Sa",1.22,205.0],["10.130.2005","250 x 120 x 200 mm yatay delikli tuğla","adet",19.0,8.2],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.012,1941.51]]],
[255,"15.220.1013","120 mm kalınlığında yatay delikli tuğla (250 x 120 x 200 mm) ile duvar yapılması","m²",782.56,847.84,"Kaba İnşaat","",[["10.130.2005","250 x 120 x 200 mm yatay delikli tuğla","adet",19.0,8.2],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.012,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.62,310.0],["10.100.1062","Düz işçi","Sa",1.24,205.0],["10.130.2002","190 x 135 x 190 mm yatay delikli tuğla","adet",26.0,6.2],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.016,1941.51],["10.130.9991","Su","m³",0.01,55.0]]],
[256,"15.220.1014","135 mm kalınlığında yatay delikli tuğla (190 x 135 x 190 mm) ile duvar yapılması","m²",808.01,875.3,"Kaba İnşaat","",[["10.130.2002","190 x 135 x 190 mm yatay delikli tuğla","adet",26.0,6.2],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.016,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.63,310.0],["10.100.1062","Düz işçi","Sa",1.26,205.0],["10.130.2002","190 x 135 x 190 mm yatay delikli tuğla","adet",36.0,6.2]]],
[257,"15.220.1015","190 mm kalınlığında yatay delikli tuğla (190 x 190 x 135 mm) ile duvar yapılması","m²",922.59,997.09,"Kaba İnşaat","",[["10.130.2002","190 x 135 x 190 mm yatay delikli tuğla","adet",36.0,6.2],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.027,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.67,310.0],["10.100.1062","Düz işçi","Sa",1.24,205.0],["10.130.2010","250 x 250 x 200 mm yatay delikli tuğla","adet",15.0,18.0],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.018,1941.51],["10.130.9991","Su","m³",0.01,55.0]]],
[258,"15.220.1016","200 mm kalınlığında yatay delikli tuğla (250 x 200 x 250 mm) ile duvar yapılması","m²",993.88,993.88,"Kaba İnşaat","",[["10.130.2010","250 x 250 x 200 mm yatay delikli tuğla","adet",15.0,18.0],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.018,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.68,310.0],["10.100.1062","Düz işçi","Sa",1.36,205.0],["10.130.2012","235 x 135 x 240 mm yatay delikli tuğla","adet",29.0,11.1]]],
[259,"15.220.1017","240 mm kalınlığında yatay delikli tuğla (235 x 240 x 135 mm) ile duvar yapılması","m²",1110.73,1110.73,"Kaba İnşaat","",[["10.130.2012","235 x 135 x 240 mm yatay delikli tuğla","adet",29.0,11.1],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.032,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",1.4,205.0],["10.130.2014","240 x 190 x 250 mm yatay delikli tuğla","adet",21.0,15.8],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.027,1941.51],["10.130.9991","Su","m³",0.01,55.0]]],
[260,"15.220.1018","250 mm kalınlığında yatay delikli tuğla (240 x 250 x 190 mm) ile duvar yapılması","m²",1119.96,1119.96,"Kaba İnşaat","",[["10.130.2014","240 x 190 x 250 mm yatay delikli tuğla","adet",21.0,15.8],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.027,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.71,310.0],["10.100.1062","Düz işçi","Sa",1.42,205.0]]],
[261,"15.220.1311","90 mm kalınlığında düşey delikli dış cephe tuğlası (190 x 90 x 50 mm) ile duvar yapılması","m²",2048.7,2048.7,"Kaba İnşaat","",[["10.130.2191","","adet",87.0,11.85],["19.100.2504","250 kg çimento dozlu harç yapılması","m³",0.024,1706.83],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",1.1,310.0],["10.100.1062","Düz işçi","Sa",1.1,205.0]]],
[262,"15.220.1312","102 mm kalınlığında düşey delikli dış cephe tuğlası (215 x 102 x 65 mm) ile duvar yapılması","m²",2549.88,2549.88,"Kaba İnşaat","",[["10.130.2193","","adet",62.0,23.15],["19.100.2504","250 kg çimento dozlu harç yapılması","m³",0.022,1706.83],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",1.1,310.0],["10.100.1062","Düz işçi","Sa",1.1,205.0],["10.130.2201","290 x 190 x 135 mm düşey delikli tuğla","adet",24.0,15.75],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.024,1941.51]]],
[263,"15.220.1411","190 mm kalınlığında düşey delikli tuğla (290 x 190 x 135 mm) ile duvar yapılması","m²",1134.44,1134.44,"Kaba İnşaat","",[["10.130.2201","290 x 190 x 135 mm düşey delikli tuğla","adet",24.0,15.75],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.024,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.67,310.0],["10.100.1062","Düz işçi","Sa",1.34,205.0],["10.130.2201","290 x 190 x 135 mm düşey delikli tuğla","adet",36.0,15.75]]],
[264,"15.220.1412","290 mm kalınlığında düşey delikli tuğla (190 x 290 x 135 mm) ile duvar yapılması","m²",1438.94,1438.94,"Kaba İnşaat","",[["10.130.2201","290 x 190 x 135 mm düşey delikli tuğla","adet",36.0,15.75],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.041,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",1.4,205.0],["10.130.2211","190 x 90 x 50 mm dolu harman tuğlası","adet",87.0,5.95],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.024,1941.51]]],
[265,"15.220.1461","90 mm kalınlığında dolu harman tuğlası (190 x 90 x 50 mm) ile duvar yapılması","m²",1246.0,1246.0,"Kaba İnşaat","",[["10.130.2211","190 x 90 x 50 mm dolu harman tuğlası","adet",87.0,5.95],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.024,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",1.2,205.0],["10.130.2212","190 x 90 x 50 mm delikli harman tuğlası","adet",87.0,5.95],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.024,1941.51]]],
[266,"15.220.1462","90 mm kalınlığında delikli harman tuğlası (190 x 90 x 50 mm) ile duvar yapılması","m²",1246.0,1246.0,"Kaba İnşaat","",[["10.130.2212","190 x 90 x 50 mm delikli harman tuğlası","adet",87.0,5.95],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.024,1941.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",1.2,205.0],["10.130.2221","","adet",13.125,22.35],["10.100.1013","Duvarcı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[267,"15.220.1612","12 cm - 13,5 cm kalınlıklar arasında teçhizatlı tuğla lentonun temini ve yerine konulması","m",1284.74,1284.74,"Kaba İnşaat","",[["10.130.2442","12 - 13,5 cm kalınlıklar arası lento tuğlası","m",1.02,897.0],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.0025,1941.51],["10.100.1013","Duvarcı ustası","Sa",0.15,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.130.2443","14,5 - 16 cm kalınlıklar arası lento tuğlası","m",1.02,937.0],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.003,1941.51],["10.100.1013","Duvarcı ustası","Sa",0.15,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0]]],
[268,"15.220.1613","14,5 cm - 16 cm kalınlıklar arasında teçhizatlı tuğla lentonun temini ve yerine konulması","m",1336.95,1336.95,"Kaba İnşaat","",[["10.130.2443","14,5 - 16 cm kalınlıklar arası lento tuğlası","m",1.02,937.0],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.003,1941.51],["10.100.1013","Duvarcı ustası","Sa",0.15,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.130.2444","18,5 - 20 cm kalınlıklar arası lento tuğlası","m",1.02,1030.0],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.004,1941.51]]],
[269,"15.220.1614","18,5 cm - 20 cm kalınlıklar arasında teçhizatlı tuğla lentonun temini ve yerine konulması","m",1502.96,1502.96,"Kaba İnşaat","",[["10.130.2444","18,5 - 20 cm kalınlıklar arası lento tuğlası","m",1.02,1030.0],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.004,1941.51],["10.100.1013","Duvarcı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0],["10.130.2445","23,5 - 25 cm kalınlıklar arası lento tuğlası","m",1.02,1130.0],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.005,1941.51],["10.100.1013","Duvarcı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[270,"15.220.1615","23,5 cm - 25 cm kalınlıklar arasında teçhizatlı tuğla lentonun temini ve yerine konulması","m",1632.89,1632.89,"Kaba İnşaat","",[["10.130.2445","23,5 - 25 cm kalınlıklar arası lento tuğlası","m",1.02,1130.0],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.005,1941.51],["10.100.1013","Duvarcı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0],["10.130.2502","","m²",1.03,173.3]]],
[271,"15.225.1013","tutkalı ile)","m²",1386.73,1386.73,"Kaba İnşaat","",[["10.130.2514","","m²",1.03,635.5],["10.100.1013","Duvarcı ustası","Sa",0.795,310.0],["10.100.1062","Düz işçi","Sa",0.795,205.0],["10.130.2515","","m²",1.03,693.3]]],
[272,"15.225.1015","tutkalı ile)","m²",1578.05,1578.05,"Kaba İnşaat","",[["10.130.2516","","m²",1.03,751.1],["10.100.1013","Duvarcı ustası","Sa",0.845,310.0],["10.100.1062","Düz işçi","Sa",0.845,205.0],["10.130.2517","","m²",1.03,809.0]]],
[273,"15.225.1063","tutkalı ile)","m²",1520.24,1520.24,"Kaba İnşaat","",[["10.130.2544","","m²",1.03,714.2],["10.100.1013","Duvarcı ustası","Sa",0.845,310.0],["10.100.1062","Düz işçi","Sa",0.845,205.0],["10.130.2545","","m²",1.03,779.1]]],
[274,"15.225.1065","tutkalı ile)","m²",1729.85,1729.85,"Kaba İnşaat","",[["10.130.2546","","m²",1.03,844.0],["10.100.1013","Duvarcı ustası","Sa",0.895,310.0],["10.100.1062","Düz işçi","Sa",0.895,205.0],["10.130.2547","","m²",1.03,909.0]]],
[275,"15.225.1113","tutkalı ile)","m²",1606.24,1606.24,"Kaba İnşaat","",[["10.130.2574","","m²",1.03,756.0],["10.100.1013","Duvarcı ustası","Sa",0.895,310.0],["10.100.1062","Düz işçi","Sa",0.895,205.0],["10.130.2575","","m²",1.03,824.7]]],
[276,"15.225.1115","tutkalı ile)","m²",1825.7,1825.7,"Kaba İnşaat","",[["10.130.2576","","m²",1.03,893.45],["10.100.1013","Duvarcı ustası","Sa",0.945,310.0],["10.100.1062","Düz işçi","Sa",0.945,205.0],["10.130.2577","","m²",1.03,962.15]]],
[277,"15.225.1151","tutkalı ile)","m²",639.65,639.65,"Kaba İnşaat","",[["10.130.2592","","m²",1.03,184.8],["10.100.1013","Duvarcı ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.130.2593","","m²",1.03,209.45]]],
[278,"15.225.1152","tutkalı ile)","m²",679.89,679.89,"Kaba İnşaat","",[["10.130.2593","","m²",1.03,209.45],["10.100.1013","Duvarcı ustası","Sa",0.61,310.0],["10.100.1062","Düz işçi","Sa",0.61,205.0]]],
[279,"15.225.1153","ile)","m²",703.19,703.19,"Kaba İnşaat","",[["10.130.2594","","m²",1.03,221.75],["10.100.1013","Duvarcı ustası","Sa",0.62,310.0],["10.100.1062","Düz işçi","Sa",0.62,205.0],["10.130.2595","","m²",1.03,246.4]]],
[280,"15.225.1154","tutkalı ile)","m²",743.43,743.43,"Kaba İnşaat","",[["10.130.2595","","m²",1.03,246.4],["10.100.1013","Duvarcı ustası","Sa",0.63,310.0],["10.100.1062","Düz işçi","Sa",0.63,205.0]]],
[281,"15.225.1155","tutkalı ile)","m²",840.78,840.78,"Kaba İnşaat","",[["10.130.2596","","m²",1.03,308.0],["10.100.1013","Duvarcı ustası","Sa",0.65,310.0],["10.100.1062","Düz işçi","Sa",0.65,205.0],["10.130.2597","","m²",1.03,332.65]]],
[282,"15.225.1156","tutkalı ile)","m²",881.01,881.01,"Kaba İnşaat","",[["10.130.2597","","m²",1.03,332.65],["10.100.1013","Duvarcı ustası","Sa",0.66,310.0],["10.100.1062","Düz işçi","Sa",0.66,205.0]]],
[283,"15.225.1157","tutkalı ile)","m²",938.11,938.11,"Kaba İnşaat","",[["10.130.2598","","m²",1.03,369.6],["10.100.1013","Duvarcı ustası","Sa",0.67,310.0],["10.100.1062","Düz işçi","Sa",0.67,205.0],["10.130.2599","","m²",1.03,431.2]]],
[284,"15.225.1158","tutkalı ile)","m²",1035.46,1035.46,"Kaba İnşaat","",[["10.130.2599","","m²",1.03,431.2],["10.100.1013","Duvarcı ustası","Sa",0.69,310.0],["10.100.1062","Düz işçi","Sa",0.69,205.0]]],
[285,"15.225.1159","tutkalı ile)","m²",1099.06,1099.06,"Kaba İnşaat","",[["10.130.2600","","m²",1.03,468.2],["10.100.1013","Duvarcı ustası","Sa",0.71,310.0],["10.100.1062","Düz işçi","Sa",0.71,205.0],["10.130.2601","","m²",1.03,492.8]]],
[286,"15.225.1160","tutkalı ile)","m²",1145.66,1145.66,"Kaba İnşaat","",[["10.130.2601","","m²",1.03,492.8],["10.100.1013","Duvarcı ustası","Sa",0.73,310.0],["10.100.1062","Düz işçi","Sa",0.73,205.0]]],
[287,"15.225.1161","tutkalı ile)","m²",1243.01,1243.01,"Kaba İnşaat","",[["10.130.2602","","m²",1.03,554.4],["10.100.1013","Duvarcı ustası","Sa",0.75,310.0],["10.100.1062","Düz işçi","Sa",0.75,205.0],["10.130.2603","","m²",1.03,616.0]]],
[288,"15.225.1162","tutkalı ile)","m²",1340.35,1340.35,"Kaba İnşaat","",[["10.130.2603","","m²",1.03,616.0],["10.100.1013","Duvarcı ustası","Sa",0.77,310.0],["10.100.1062","Düz işçi","Sa",0.77,205.0]]],
[289,"15.225.1163","tutkalı ile)","m²",1440.93,1440.93,"Kaba İnşaat","",[["10.130.2604","","m²",1.03,677.6],["10.100.1013","Duvarcı ustası","Sa",0.795,310.0],["10.100.1062","Düz işçi","Sa",0.795,205.0],["10.130.2605","","m²",1.03,739.2]]],
[290,"15.225.1164","tutkalı ile)","m²",1541.48,1541.48,"Kaba İnşaat","",[["10.130.2605","","m²",1.03,739.2],["10.100.1013","Duvarcı ustası","Sa",0.82,310.0],["10.100.1062","Düz işçi","Sa",0.82,205.0]]],
[291,"15.225.1165","tutkalı ile)","m²",1642.04,1642.04,"Kaba İnşaat","",[["10.130.2606","","m²",1.03,800.8],["10.100.1013","Duvarcı ustası","Sa",0.845,310.0],["10.100.1062","Düz işçi","Sa",0.845,205.0],["10.130.2607","","m²",1.03,862.4]]],
[292,"15.225.1166","tutkalı ile)","m²",1742.59,1742.59,"Kaba İnşaat","",[["10.130.2607","","m²",1.03,862.4],["10.100.1013","Duvarcı ustası","Sa",0.87,310.0],["10.100.1062","Düz işçi","Sa",0.87,205.0]]],
[293,"15.230.1301","10 cm kalınlığındaki teçhizatlı bimsbeton lento temini ve yerine konulması","m²",790.79,790.79,"Kaba İnşaat","ÖLÇÜ: Projesindeki boyutlar üzerinden hesaplanır.",[["10.130.2951","10 cm kalınlıkta techizatlı bimsbeton lento","m²",1.03,416.0],["10.100.1013","Duvarcı ustası","Sa",0.28,310.0],["10.100.1062","Düz işçi","Sa",0.56,205.0],["10.130.2952","13,5 cm kalınlıkta techizatlı bimsbeton lento","m²",1.03,583.0],["10.100.1013","Duvarcı ustası","Sa",0.31,310.0],["10.100.1062","Düz işçi","Sa",0.62,205.0]]],
[294,"15.230.1302","13,5 cm kalınlığındaki teçhizatlı bimsbeton lento temini ve yerine konulması","m²",1033.93,1033.93,"Kaba İnşaat","ÖLÇÜ: Projesindeki boyutlar üzerinden hesaplanır.",[["10.130.2952","13,5 cm kalınlıkta techizatlı bimsbeton lento","m²",1.03,583.0],["10.100.1013","Duvarcı ustası","Sa",0.31,310.0],["10.100.1062","Düz işçi","Sa",0.62,205.0],["10.130.2953","15 cm kalınlıkta techizatlı bimsbeton lento","m²",1.03,652.0],["10.100.1013","Duvarcı ustası","Sa",0.33,310.0],["10.100.1062","Düz işçi","Sa",0.66,205.0]]],
[295,"15.230.1303","15 cm kalınlığındaki teçhizatlı bimsbeton lento temini ve yerine konulması","m²",1141.26,1141.26,"Kaba İnşaat","ÖLÇÜ: Projesindeki boyutlar üzerinden hesaplanır.",[["10.130.2953","15 cm kalınlıkta techizatlı bimsbeton lento","m²",1.03,652.0],["10.100.1013","Duvarcı ustası","Sa",0.33,310.0],["10.100.1062","Düz işçi","Sa",0.66,205.0],["10.130.2954","19 cm kalınlıkta techizatlı bimsbeton lento","m²",1.03,836.0],["10.100.1013","Duvarcı ustası","Sa",0.37,310.0],["10.100.1062","Düz işçi","Sa",0.74,205.0]]],
[296,"15.230.1304","19 cm kalınlığındaki teçhizatlı bimsbeton lento temini ve yerine konulması","m²",1415.44,1415.44,"Kaba İnşaat","ÖLÇÜ: Projesindeki boyutlar üzerinden hesaplanır.",[["10.130.2954","19 cm kalınlıkta techizatlı bimsbeton lento","m²",1.03,836.0],["10.100.1013","Duvarcı ustası","Sa",0.37,310.0],["10.100.1062","Düz işçi","Sa",0.74,205.0],["10.130.3101","14 cm toplam kalınlıkta duvar bloğu","m²",1.03,361.0]]],
[297,"15.235.1051","15 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",676.59,676.59,"Kaba İnşaat","",[["10.130.3501","","m³",0.158,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.8,205.0],["10.130.3501","","m³",0.18,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.45,310.0],["10.100.1062","Düz işçi","Sa",0.9,205.0]]],
[298,"15.235.1052","17,5 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",765.68,765.68,"Kaba İnşaat","",[["10.130.3501","","m³",0.18,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.45,310.0],["10.100.1062","Düz işçi","Sa",0.9,205.0],["10.130.3501","","m³",0.21,1603.0]]],
[299,"15.235.1053","20 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",870.79,870.79,"Kaba İnşaat","",[["10.130.3501","","m³",0.21,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.3501","","m³",0.236,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.55,310.0],["10.100.1062","Düz işçi","Sa",1.1,205.0]]],
[300,"15.235.1054","22,5 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",967.89,967.89,"Kaba İnşaat","",[["10.130.3501","","m³",0.236,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.55,310.0],["10.100.1062","Düz işçi","Sa",1.1,205.0],["10.130.3501","","m³",0.263,1603.0]]],
[301,"15.235.1055","25 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",1066.99,1066.99,"Kaba İnşaat","",[["10.130.3501","","m³",0.263,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",1.2,205.0],["10.130.3501","","m³",0.289,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.65,310.0],["10.100.1062","Düz işçi","Sa",1.3,205.0]]],
[302,"15.235.1056","27,5 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",1164.09,1164.09,"Kaba İnşaat","",[["10.130.3501","","m³",0.289,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.65,310.0],["10.100.1062","Düz işçi","Sa",1.3,205.0],["10.130.3501","","m³",0.315,1603.0]]],
[303,"15.235.1057","30 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",1261.19,1261.19,"Kaba İnşaat","",[["10.130.3501","","m³",0.315,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",1.4,205.0],["10.130.3501","","m³",0.341,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.75,310.0],["10.100.1062","Düz işçi","Sa",1.5,205.0]]],
[304,"15.235.1058","32,5 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",1358.28,1358.28,"Kaba İnşaat","",[["10.130.3501","","m³",0.341,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.75,310.0],["10.100.1062","Düz işçi","Sa",1.5,205.0],["10.130.3501","","m³",0.368,1603.0]]],
[305,"15.235.1059","35 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",1457.38,1457.38,"Kaba İnşaat","",[["10.130.3501","","m³",0.368,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.130.3501","","m³",0.394,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.85,310.0],["10.100.1062","Düz işçi","Sa",1.7,205.0]]],
[306,"15.235.1060","37,5 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",1554.48,1554.48,"Kaba İnşaat","",[["10.130.3501","","m³",0.394,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.85,310.0],["10.100.1062","Düz işçi","Sa",1.7,205.0],["10.130.3501","","m³",0.42,1603.0]]],
[307,"15.235.1061","40 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması","m²",1651.58,1651.58,"Kaba İnşaat","",[["10.130.3501","","m³",0.42,1603.0],["10.100.1013","Duvarcı ustası","Sa",0.9,310.0],["10.100.1062","Düz işçi","Sa",1.8,205.0],["10.130.3201","","adet",14.4,3.52]]],
[308,"15.245.1001","150 gr/m² ağırlıkta geotekstil keçe serilmesi","m²",52.75,52.75,"Kaba İnşaat","",[["10.330.6002","Geotekstil keçe (150 gr/m2)","m²",1.1,7.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["10.330.6004","Geotekstil keçe (250 gr/m2)","m²",1.1,11.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[309,"15.245.1002","250 gr/m² ağırlıkta geotekstil keçe serilmesi","m²",58.25,58.25,"Kaba İnşaat","",[["10.330.6004","Geotekstil keçe (250 gr/m2)","m²",1.1,11.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["10.330.6007","Geotekstil keçe (500 gr/m2)","m²",1.1,22.0]]],
[310,"15.245.1003","500 gr/m² ağırlıkta geotekstil keçe serilmesi","m²",73.38,73.38,"Kaba İnşaat","",[["10.330.6007","Geotekstil keçe (500 gr/m2)","m²",1.1,22.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["19.100.2501","200 Kg çimento dozlu harç yapılması","m³",0.035,1779.65],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[311,"15.250.1011","200 kg çimento dozlu tesviye tabakası yapılması","m²",335.74,335.74,"İnce İnşaat","",[["19.100.2501","200 Kg çimento dozlu harç yapılması","m³",0.035,1779.65],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["19.100.2511","400 dozlu ince harç","m³",0.025,2182.8],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.001,9000.0]]],
[312,"15.250.1111","2.5 cm kalınlığında 400 kg çimento dozlu şap yapılması","m²",449.53,488.76,"İnce İnşaat","",[["19.100.2511","400 dozlu ince harç","m³",0.025,2182.8],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.001,9000.0],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.35,310.0],["10.100.1041","Marangoz usta yardımcısı","Sa",0.1,230.0],["10.100.1062","Düz işçi","Sa",0.7,205.0],["10.100.1062","Düz işçi","Sa",0.1,205.0],["19.100.2512","450 Kg çimento dozlu harç yapılması","m³",0.025,2220.13]]],
[313,"15.250.1112","2.5 cm kalınlığında 450 kg çimento dozlu şap yapılması","m²",450.69,450.69,"İnce İnşaat","",[["19.100.2512","450 Kg çimento dozlu harç yapılması","m³",0.025,2220.13],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.001,9000.0],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.35,310.0],["10.100.1041","Marangoz usta yardımcısı","Sa",0.1,230.0],["10.100.1062","Düz işçi","Sa",0.7,205.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[314,"15.250.1113","2.5 cm kalınlığında 500 kg çimento dozlu şap yapılması","m²",463.95,463.95,"İnce İnşaat","",[["19.100.2515","500 Kg çimento dozlu ince harç yapılması","m³",0.025,2644.45],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.001,9000.0],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.35,310.0],["10.100.1041","Marangoz usta yardımcısı","Sa",0.1,230.0],["10.100.1062","Düz işçi","Sa",0.7,205.0],["10.100.1062","Düz işçi","Sa",0.1,205.0],["19.100.1100","Sıva Makinası","Sa",0.1,749.85]]],
[315,"15.250.1114","Makina ile ortalama 2,5 cm kalınlıkta alçı esaslı şap yapılması","m²",385.59,385.59,"İnce İnşaat","",[["19.100.1100","Sıva Makinası","Sa",0.1,749.85],["10.130.9991","Su","m³",0.025,55.0],["10.100.1068","Birinci sınıf usta","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[316,"15.255.1001","kalınlıkta plastomer esaslı (-5 °C soğukta bükülmeli) polyester keçe taşıyıcılı polimer","m²",586.01,946.95,"Kaba İnşaat","",[["10.330.5192","Plastomer esaslı cam tülü taşıyıcılı örtü","m²",1.15,112.0],["10.330.5201","Plastomer esaslı polyester keçe taşıyıcılı örtü","m²",1.15,130.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[317,"15.255.1002","kalınlıkta plastomer esaslı (-10 °C soğukta bükülmeli) polyester keçe taşıyıcılı polimer","m²",611.89,611.89,"Kaba İnşaat","",[["10.330.5102","Plastomer esaslı cam tülü taşıyıcılı örtü","m²",1.15,120.0],["10.330.5121","Plastomer esaslı polyester keçe taşıyıcılı örtü","m²",1.15,140.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[318,"15.255.1003","kalınlıkta elastomer esaslı (-20 °C soğukta bükülmeli) polyester keçe taşıyıcılı polimer","m²",679.45,679.45,"Kaba İnşaat","",[["10.330.5152","Elastomer esaslı cam tülü taşıyıcılı örtü","m²",1.15,137.0],["10.330.5171","Elastomer polyester keçe taşıyıcılı örtü","m²",1.15,170.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[319,"15.255.1010","kalınlıkta plastomer esaslı (-5 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü","m²",629.14,629.14,"Kaba İnşaat","",[["10.330.5192","Plastomer esaslı cam tülü taşıyıcılı örtü","m²",1.15,112.0],["10.330.5202","","m²",1.15,160.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[320,"15.255.1011","kalınlıkta plastomer esaslı (-10 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü","m²",655.01,655.01,"Kaba İnşaat","",[["10.330.5102","Plastomer esaslı cam tülü taşıyıcılı örtü","m²",1.15,120.0],["10.330.5122","","m²",1.15,170.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[321,"15.255.1012","kalınlıkta elastomer esaslı (-20 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü","m²",715.39,715.39,"Kaba İnşaat","",[["10.330.5152","Elastomer esaslı cam tülü taşıyıcılı örtü","m²",1.15,137.0],["10.330.5172","","m²",1.15,195.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[322,"15.255.1013","mm kalınlıkta plastomer esaslı (-5 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü","m²",655.01,655.01,"Kaba İnşaat","",[["10.330.5201","Plastomer esaslı polyester keçe taşıyıcılı örtü","m²",1.15,130.0],["10.330.5202","","m²",1.15,160.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[323,"15.255.1014","mm kalınlıkta plastomer esaslı (-10 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü","m²",683.76,683.76,"Kaba İnşaat","",[["10.330.5121","Plastomer esaslı polyester keçe taşıyıcılı örtü","m²",1.15,140.0],["10.330.5122","","m²",1.15,170.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[324,"15.255.1015","mm kalınlıkta elastomer esaslı (-20 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü","m²",762.83,762.83,"Kaba İnşaat","",[["10.330.5171","Elastomer polyester keçe taşıyıcılı örtü","m²",1.15,170.0],["10.330.5172","","m²",1.15,195.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[325,"15.260.1003","1,5 mm kalınlıkta, PVC esaslı, (UV dayanımlı, donatılı) jeomembran ile su yalıtımı yapılması","m²",456.65,456.65,"Kaba İnşaat","",[["10.330.6022","","m²",1.05,190.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.330.6023","","m²",1.05,250.0]]],
[326,"15.260.1004","2 mm kalınlıkta, PVC esaslı, (UV dayanımlı, donatılı) jeomembran ile su yalıtımı yapılması","m²",535.4,535.4,"Kaba İnşaat","",[["10.330.6023","","m²",1.05,250.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[327,"15.260.1008","2 mm kalınlıkta, HDPE esaslı, (UV dayanımlı, donatılı) jeomembran ile su yalıtımı yapılması","m²",482.9,482.9,"Kaba İnşaat","",[["10.330.6043","","m²",1.05,210.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[328,"15.260.1011","1,5 mm kalınlıkta, Termoset EPDM esaslı, jeomembran ile su yalıtımı yapılması","m²",692.9,692.9,"Kaba İnşaat","",[["10.330.6062","","m²",1.05,370.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.330.6063","","m²",1.05,510.0]]],
[329,"15.260.1012","2 mm kalınlıkta, Termoset EPDM esaslı, jeomembran ile su yalıtımı yapılması","m²",876.65,876.65,"Kaba İnşaat","",[["10.330.6063","","m²",1.05,510.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[330,"15.260.1013","1,5 mm kalınlıkta, TPO esaslı, (UV dayanımlı, donatılı) jeomembran ile su yalıtımı yapılması","m²",541.96,541.96,"Kaba İnşaat","",[["10.330.6072","","m²",1.05,255.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.330.6073","","m²",1.05,340.0]]],
[331,"15.260.1014","2 mm kalınlıkta, TPO esaslı, (UV dayanımlı, donatılı) jeomembran ile su yalıtımı yapılması","m²",653.53,653.53,"Kaba İnşaat","",[["10.330.6073","","m²",1.05,340.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[332,"15.265.1001","3 mm kalınlıkta, HDPE levhalarla, su yalıtımı yapılması","m²",603.21,603.21,"Kaba İnşaat","",[["10.330.6302","3 mm kalınlıkta HDPE levha","m²",1.05,290.0],["10.420.1012","Vida ve plastik dübel","adet",3.0,2.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[333,"15.265.1002","4 mm kalınlıkta, HDPE levhalarla, su yalıtımı yapılması","m²",736.03,736.03,"Kaba İnşaat","",[["10.330.6303","4 mm kalınlıkta HDPE levha","m²",1.05,390.0],["10.420.1012","Vida ve plastik dübel","adet",3.0,2.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[334,"15.265.1003","5 mm kalınlıkta, HDPE levhalarla, su yalıtımı yapılması","m²",862.28,862.28,"Kaba İnşaat","",[["10.330.6304","5 mm kalınlıkta HDPE levha","m²",1.05,485.0],["10.420.1012","Vida ve plastik dübel","adet",3.0,2.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[335,"15.265.1004","3 mm kalınlıkta, PP levhalarla, su yalıtımı yapılması","m²",577.03,577.03,"Kaba İnşaat","",[["10.330.6322","3 mm kalınlıkta PP levha","m²",1.05,270.0],["10.420.1012","Vida ve plastik dübel","adet",3.0,2.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[336,"15.265.1005","4 mm kalınlıkta, PP levhalarla, su yalıtımı yapılması","m²",709.85,709.85,"Kaba İnşaat","",[["10.330.6323","4 mm kalınlıkta PP levha","m²",1.05,370.0],["10.420.1012","Vida ve plastik dübel","adet",3.0,2.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[337,"15.265.1006","5 mm kalınlıkta, PP levhalarla, su yalıtımı yapılması","m²",849.24,849.24,"Kaba İnşaat","",[["10.330.6324","5 mm kalınlıkta PP levha","m²",1.05,475.0],["10.420.1012","Vida ve plastik dübel","adet",3.0,2.0],["10.100.1010","Yalıtımcı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[338,"15.270.1202","Örgülü Geotekstil, Üst Tabaka 200 gr/m² PP Örgüsüz Geotekstil, Toplam Ağırlık 5500","m²",293.75,293.75,"Kaba İnşaat","",[["10.450.5152","","m²",1.1,120.0],["10.100.1010","Yalıtımcı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.450.5153","","m²",1.1,125.0]]],
[339,"15.270.1203","Örgülü Geotekstil, Üst Tabaka 200 gr/m² PP Örgüsüz Geotekstil, Toplam Ağırlık 6500","m²",300.63,300.63,"Kaba İnşaat","",[["10.450.5153","","m²",1.1,125.0],["10.100.1010","Yalıtımcı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[340,"15.275.1011","Taş duvar yüzeylerine gömme oluklu derz yapılması","m²",265.98,265.98,"Kaba İnşaat","",[["19.100.2514","500 Kg çimento dozlu harç yapılması","m³",0.01,2348.45],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.2,310.0],["10.100.1045","Duvarcı usta yardımcısı","Sa",0.15,230.0],["10.100.1062","Düz işçi","Sa",0.35,205.0],["10.100.1062","Düz işçi","Sa",0.1,205.0],["19.100.2514","500 Kg çimento dozlu harç yapılması","m³",0.015,2348.45],["10.130.9991","Su","m³",0.01,55.0]]],
[341,"15.275.1012","Taş duvar yüzeylerine kabartma derz yapılması","m²",290.35,290.35,"Kaba İnşaat","",[["19.100.2514","500 Kg çimento dozlu harç yapılması","m³",0.015,2348.45],["10.130.9991","Su","m³",0.01,55.0],["10.100.1013","Duvarcı ustası","Sa",0.25,310.0],["10.100.1045","Duvarcı usta yardımcısı","Sa",0.25,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.100.1062","Düz işçi","Sa",0.1,205.0],["19.100.2504","250 kg çimento dozlu harç yapılması","m³",0.023,1706.83],["19.100.2508","350 Kg çimento dozlu ince harç yapılması","m³",0.01,2054.48]]],
[342,"15.275.1111","250/350 kg çimento dozlu kaba ve ince harçla sıva yapılması (dış cephe sıvası)","m²",744.5,811.56,"Kaba İnşaat","",[["19.100.2504","250 kg çimento dozlu harç yapılması","m³",0.023,1706.83],["19.100.2508","350 Kg çimento dozlu ince harç yapılması","m³",0.01,2054.48],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",1.0,310.0],["10.100.1044","Sıvacı usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.4,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["19.100.2522","","m³",0.023,2077.5]]],
[343,"15.275.1112","200/250 kg kireç/çimento karışımı kaba ve ince harçla sıva yapılması (iç cephe sıvası)","m²",661.76,721.5,"Kaba İnşaat","",[["19.100.2522","","m³",0.023,2077.5],["19.100.2521","","m³",0.01,2032.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.9,310.0],["10.100.1044","Sıvacı usta yardımcısı","Sa",0.3,230.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[344,"15.275.1113","250/350 kg kireç/çimento karışımı kaba ve ince harçla sıva yapılması (tavan sıvası)","m²",693.65,693.65,"Kaba İnşaat","",[["19.100.2507","350 Kg çimento dozlu harç yapılması","m³",0.01,1963.48],["19.100.2521","","m³",0.015,2032.51],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.9,310.0],["10.100.1044","Sıvacı usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.4,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.023,1941.51]]],
[345,"15.275.1114","250/350 kg çimento dozlu kaba ve ince harçla serpme (çarpma) sıva yapılması","m²",517.38,517.38,"Kaba İnşaat","",[["19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",0.023,1941.51],["19.100.2507","350 Kg çimento dozlu harç yapılması","m³",0.015,1963.48],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.5,310.0],["10.100.1044","Sıvacı usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["19.100.2508","350 Kg çimento dozlu ince harç yapılması","m³",0.02,2054.48]]],
[346,"15.275.1115","350 kg çimento dozlu harçla tek kat ince sıva yapılması","m²",483.61,483.61,"Kaba İnşaat","",[["19.100.2508","350 Kg çimento dozlu ince harç yapılması","m³",0.02,2054.48],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.75,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["19.100.2504","250 kg çimento dozlu harç yapılması","m³",0.023,1706.83],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.5,310.0]]],
[347,"15.275.1116","250 kg çimento dozlu harç ile kaba sıva yapılması","m²",443.51,443.51,"Kaba İnşaat","",[["19.100.2504","250 kg çimento dozlu harç yapılması","m³",0.023,1706.83],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.5,310.0],["10.100.1044","Sıvacı usta yardımcısı","Sa",0.25,230.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["19.100.2522","","m³",0.023,2077.5]]],
[348,"15.275.1117","200 kg Çimento ve Kireç karşımı harç ile kaba sıva yapılması (iç cephe)","m²",454.16,454.16,"Kaba İnşaat","",[["19.100.2522","","m³",0.023,2077.5],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.5,310.0],["10.100.1044","Sıvacı usta yardımcısı","Sa",0.25,230.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["19.100.2505","300 Kg çimento dozlu harç yapılması","m³",0.03,1835.15],["10.130.9991","Su","m³",0.01,55.0]]],
[349,"15.275.9992","Metal kapı kasa arkalarının beton harcı ile doldurulması","m²",623.56,623.56,"Kaba İnşaat","",[["19.100.2505","300 Kg çimento dozlu harç yapılması","m³",0.03,1835.15],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",0.65,205.0],["19.100.2433","Perlitli alçı harcı hazırlanması","m³",0.015,2219.73],["10.130.9991","Su","m³",0.002,55.0],["19.100.2433","Perlitli alçı harcı hazırlanması","m³",0.0025,2219.73],["19.100.2432","Saten alçı harcı hazırlanması","m³",0.0025,4243.0]]],
[350,"15.280.1009","Perlitli sıva alçısı ve saten alçı ile kaplama yapılması (Beton, tuğla duvar vb. yüzeylere)","m²",694.59,759.29,"İnce İnşaat","1) Projedeki ölçülere göre, sıvanan bütün yüzler (boşluk yanları dâhil) hesaplanır.",[["19.100.2433","Perlitli alçı harcı hazırlanması","m³",0.015,2219.73],["10.130.9991","Su","m³",0.002,55.0],["19.100.2433","Perlitli alçı harcı hazırlanması","m³",0.0025,2219.73],["19.100.2432","Saten alçı harcı hazırlanması","m³",0.0025,4243.0],["10.130.9991","Su","m³",0.005,55.0],["10.200.3141","Alçı sıva köşe profili ≥ 0.40 mm kalınlıkta","m",0.1,5.45],["10.330.2502","Sıva filesi","m²",0.2,10.0],["19.100.2432","Saten alçı harcı hazırlanması","m³",0.001,4243.0]]],
[351,"15.280.1011","Saten alçı kaplaması yapılması (ortalama 1 mm kalınlık)","m²",139.71,139.71,"İnce İnşaat","",[["19.100.2432","Saten alçı harcı hazırlanması","m³",0.001,4243.0],["10.130.9991","Su","m³",0.005,55.0],["10.300.1602","Zımpara kağıdı","adet",0.5,8.5],["10.100.1012","Sıvacı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[352,"15.280.1012","Makina alçısı ile tavanlara 15 mm kalınlığında tek kat alçı sıva yapılması","m²",468.55,468.55,"İnce İnşaat","",[["19.100.1100","Sıva Makinası","Sa",0.06,749.85],["10.200.3141","Alçı sıva köşe profili ≥ 0.40 mm kalınlıkta","m",0.1,5.45],["10.130.9991","Su","m³",0.01,55.0],["10.100.1012","Sıvacı ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.35,205.0]]],
[353,"15.300.1001","Ahşaptan oturtma çatı yapılması (çatı örtüsü altı tahta kaplamalı)","m²",1417.89,1541.94,"Kaba İnşaat","1) Çatının onanmış projesinden yatay düzlemdeki izdüşümü saçak dışından saçak dışına (oluk hariç) ölçülerek m² olarak",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.05,9000.0],["19.100.1091","","Sa",0.02,6059.7],["10.100.1017","Dülger ustası","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[354,"15.300.1002","Ahşaptan oturtma çatı yapılması (çatı örtüsünün altı OSB/3 kaplamalı)","m²",1566.83,1566.83,"Kaba İnşaat","1) Çatının onanmış projesinden yatay düzlemdeki izdüşümü saçak dışından saçak dışına (oluk hariç) ölçülerek m² olarak",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.03,9000.0],["10.170.1925","","m²",1.1,305.0],["19.100.1091","","Sa",0.014,6059.7],["10.100.1017","Dülger ustası","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[355,"15.300.1003","Ahşaptan makaslı çatı yapılması","m³",28881.34,28881.34,"Kaba İnşaat","Onanmış projesindeki boyutlara göre çatıda kullanılan kereste miktarı hesaplanır. Ahşap aksamın boyutları bu aksamın içi",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",1.1,9000.0],["19.100.1091","","Sa",0.4,6059.7],["10.100.1017","Dülger ustası","Sa",20.0,310.0],["10.100.1062","Düz işçi","Sa",20.0,205.0]]],
[356,"15.300.1004","Rendeli ahşaptan makaslı çatı yapılması","m³",29746.83,29746.83,"Kaba İnşaat","Onanmış projesindeki boyutlara göre çatıda kullanılan kereste miktarı hesaplanır. Ahşap aksamın boyutları bu aksamın içi",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",1.15,9000.0],["19.100.1091","","Sa",0.44,6059.7],["10.100.1017","Dülger ustası","Sa",20.0,310.0],["10.100.1062","Düz işçi","Sa",20.0,205.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.024,9000.0]]],
[357,"15.300.1005","Çatı üzerine tahta kaplama yapılması","m²",611.06,611.06,"Kaba İnşaat","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.024,9000.0],["19.100.1091","","Sa",0.01,6059.7],["10.100.1017","Dülger ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[358,"15.300.1006","Çatı üzerine OSB/3 kaplama yapılması","m²",684.69,684.69,"Kaba İnşaat","",[["10.170.1925","","m²",1.1,305.0],["10.100.1017","Dülger ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.04,9000.0],["19.100.1091","","Sa",0.012,6059.7]]],
[359,"15.300.1007","Rendeli ahşaptan saçak altı ve alın kaplaması yapılması","m²",1031.53,1031.53,"Kaba İnşaat","Proje üzerinde tahta kaplı saçağın yatay düzlemdeki izdüşümü ölçülür.",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.04,9000.0],["19.100.1091","","Sa",0.012,6059.7],["10.100.1017","Dülger ustası","Sa",0.75,310.0],["10.100.1062","Düz işçi","Sa",0.75,205.0]]],
[360,"15.305.1201","Renksiz beton kiremitler ile çatı örtüsü yapılması (çift latalı sistem)","m²",949.75,949.75,"Kaba İnşaat","Eğimli alanlar projesi üzerinden hesaplanır. 0,10 m² ve daha küçük boşluklar düşülmez.",[["10.130.4101","Beton kiremit (renksiz)","m²",1.05,144.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.01,9000.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[361,"15.305.1202","Demir oksit boyalı beton kiremitler ile çatı örtüsü yapılması (çift latalı sistem)","m²",994.38,994.38,"Kaba İnşaat","Eğimli alanlar projesi üzerinden hesaplanır. 0,10 m² ve daha küçük boşluklar düşülmez.",[["10.130.4103","Beton kiremit (demir oksit boyalı)","m²",1.05,178.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.01,9000.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[362,"15.305.1204","Renksiz beton mahya kiremitleri ile mahya yapılması","m",567.38,567.38,"Kaba İnşaat","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.0025,9000.0],["10.130.4202","","m",1.05,155.0],["10.130.4102","Beton mahya kiremiti (renksiz)","m",1.05,91.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[363,"15.305.1205","Demir oksit boyalı beton mahya kiremitleri ile mahya yapılması","m",604.13,604.13,"Kaba İnşaat","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.0025,9000.0],["10.130.4202","","m",1.05,155.0],["10.130.4104","Beton mahya kiremiti (demir oksit boyalı)","m",1.05,119.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[364,"15.305.1206","Demir oksit boyalı üzeri renkli sırla kaplanmış beton mahya kiremitleri ile mahya yapılması","m",644.81,644.81,"Kaba İnşaat","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.0025,9000.0],["10.130.4202","","m",1.05,155.0],["10.130.4106","","m",1.05,150.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[365,"15.305.1207","Renksiz perlitli beton kiremitler ile çatı örtüsü yapılması (çift latalı sistem)","m²",910.38,910.38,"Kaba İnşaat","",[["10.130.4121","Perlitli beton kiremit (renksiz)","m²",1.05,114.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.01,9000.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[366,"15.305.1208","Demir oksit boyalı perlitli beton kiremitler ile çatı örtüsü yapılması (çift latalı sistem)","m²",949.75,949.75,"Kaba İnşaat","",[["10.130.4123","Perlitli beton kiremit (demir oksit boyalı)","m²",1.05,144.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.01,9000.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[367,"15.305.1210","Renksiz perlitli beton mahya kiremitleri ile mahya yapılması","m",552.94,552.94,"Kaba İnşaat","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.0025,9000.0],["10.130.4202","","m",1.05,155.0],["10.130.4122","Perlitli beton mahya kiremiti (renksiz)","m",1.05,80.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[368,"15.305.1211","Demir oksit boyalı perlitli beton mahya kiremitleri ile mahya yapılması","m",568.69,568.69,"Kaba İnşaat","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.0025,9000.0],["10.130.4202","","m",1.05,155.0],["10.130.4124","","m",1.05,92.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[369,"15.305.1212","Demir oksit boyalı üzeri renkli sırla kaplanmış beton mahya kiremitleri ile mahya yapılması","m",623.81,623.81,"Kaba İnşaat","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.0025,9000.0],["10.130.4202","","m",1.05,155.0],["10.130.4126","","m",1.05,134.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[370,"15.305.1213","yapışkanlı, UV dayanımlı duvar baca dibi bandı ile duvar, baca dibi vb. yerlerde su yalıtımı","m",472.31,472.31,"Kaba İnşaat","",[["10.130.4204","","m",1.1,306.0],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",0.05,205.0],["10.130.4206","","m",1.05,57.5],["10.420.1012","Vida ve plastik dübel","adet",4.0,2.0]]],
[371,"15.305.1214","Alüminyum baskı çıtası ve poliüretan mastik ile yalıtım bitişlerinde sızdırmazlık sağlanması","m",294.83,294.83,"Kaba İnşaat","",[["10.130.4206","","m",1.05,57.5],["10.420.1012","Vida ve plastik dübel","adet",4.0,2.0],["10.300.2157","","adet",0.2,240.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.100.1016","Kiremit tipi çatı kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[372,"15.310.1001","12 no'lu çinko levhadan 150 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.","m",1085.9,1085.9,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.75,310.0],["10.100.1064","Çırak","Sa",0.75,205.0]]],
[373,"15.310.1002","12 no'lu çinko levhadan 120 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.","m",973.39,973.39,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.75,310.0],["10.100.1064","Çırak","Sa",0.75,205.0]]],
[374,"15.310.1003","12 no'lu çinko levhadan 100 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.","m",898.98,898.98,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.75,310.0],["10.100.1064","Çırak","Sa",0.75,205.0]]],
[375,"15.310.1004","10 no'lu çinko levhadan 100 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.","m",820.96,820.96,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.75,310.0],["10.100.1064","Çırak","Sa",0.75,205.0]]],
[376,"15.310.1005","10 no'lu çinko levhadan 80 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.","m",776.16,776.16,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.75,310.0],["10.100.1064","Çırak","Sa",0.75,205.0]]],
[377,"15.310.1006","12 no'lu çinko levhadan 80 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.","m",842.91,842.91,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.75,310.0],["10.100.1064","Çırak","Sa",0.75,205.0]]],
[378,"15.310.1007","10 no'lu çinko levhadan 75 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.","m",747.55,747.55,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.75,310.0],["10.100.1064","Çırak","Sa",0.75,205.0]]],
[379,"15.310.1008","10 no'lu çinko levhadan 70 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.","m",707.86,707.86,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.75,310.0],["10.100.1064","Çırak","Sa",0.75,205.0]]],
[380,"15.310.1101","14 no'lu çinko levhadan 240 mm çapında yağmur oluğu yapılması ve yerine tespiti","m",2264.46,2264.46,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",0.3,6.0],["10.100.1026","Tenekeci ustası","Sa",1.5,310.0],["10.100.1064","Çırak","Sa",1.5,205.0]]],
[381,"15.310.1102","12 no'lu çinko levhadan 185 mm çapında yağmur oluğu yapılması ve yerine tespiti.","m",1877.18,1877.18,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",0.3,6.0],["10.100.1026","Tenekeci ustası","Sa",1.5,310.0],["10.100.1064","Çırak","Sa",1.5,205.0]]],
[382,"15.310.1103","12 no'lu çinko levhadan 155 mm çapında yağmur oluğu yapılması ve yerine tespiti.","m",1748.6,1748.6,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",0.3,6.0],["10.100.1026","Tenekeci ustası","Sa",1.5,310.0],["10.100.1064","Çırak","Sa",1.5,205.0]]],
[383,"15.310.1104","12 no'lu çinko levhadan 130 mm çapında yağmur oluğu yapılması ve yerine tespiti.","m",1608.45,1608.45,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",0.3,6.0],["10.100.1026","Tenekeci ustası","Sa",1.5,310.0],["10.100.1064","Çırak","Sa",1.5,205.0]]],
[384,"15.310.1105","12 no'lu çinko levhadan 110 mm çapında yağmur oluğu yapılması ve yerine tespiti.","m",1534.06,1534.06,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",0.3,6.0],["10.100.1026","Tenekeci ustası","Sa",1.5,310.0],["10.100.1064","Çırak","Sa",1.5,205.0]]],
[385,"15.310.1106","12 no'lu çinko levhadan 90 mm çapında yağmur oluğu yapılması ve yerine tespiti.","m",1438.16,1438.16,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",0.3,6.0],["10.100.1026","Tenekeci ustası","Sa",1.5,310.0],["10.100.1064","Çırak","Sa",1.5,205.0]]],
[386,"15.310.1201","14 no.lu çinko levhadan eğimli çatı deresi yapılması ve yerine konulması","m",1344.89,1344.89,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",0.9,6.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.0009,9000.0],["10.100.1026","Tenekeci ustası","Sa",0.8,310.0],["10.100.1064","Çırak","Sa",0.8,205.0]]],
[387,"15.310.1202","14 no.lu çinko levhadan oluk halinde yatay çatı deresi yapılması ve yerine konulması","m",2626.0,2626.0,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",1.5,6.0],["10.100.1026","Tenekeci ustası","Sa",1.8,310.0],["10.100.1064","Çırak","Sa",1.8,205.0]]],
[388,"15.310.1204","14 no.lu çinko levhadan atika duvar arkası için çatı deresi yapılması ve yerine konulması","m",2818.48,2818.48,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",1.5,6.0],["10.100.1026","Tenekeci ustası","Sa",1.75,310.0],["10.100.1064","Çırak","Sa",1.75,205.0]]],
[389,"15.310.1207","12 no.lu çinko levhadan pencere denizliği yapılması ve yerine konulması","m",981.49,981.49,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",0.45,6.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.0004,9000.0],["10.100.1026","Tenekeci ustası","Sa",0.85,310.0],["10.100.1064","Çırak","Sa",0.85,205.0]]],
[390,"15.310.1208","12 no.lu çinko levhadan baca temizleme kutusu yapılması ve yerine konulması","adet",575.98,575.98,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.5,310.0],["10.100.1064","Çırak","Sa",0.5,205.0]]],
[391,"15.310.1209","12 no.lu çinko levhadan soba deliği ağızlığı ve kapağı yapılması ve yerine konulması","adet",412.8,412.8,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.35,310.0],["10.100.1064","Çırak","Sa",0.35,205.0],["10.100.1026","Tenekeci ustası","Sa",0.9,310.0],["10.100.1064","Çırak","Sa",0.9,205.0]]],
[392,"15.310.1301","0.50 mm bakır levhadan 125 mm çapında düşey yağmur borusu yapılması ve yerine tespiti","m",1955.96,1955.96,"Kaba İnşaat","",[["10.100.1026","Tenekeci ustası","Sa",0.9,310.0],["10.100.1064","Çırak","Sa",0.9,205.0]]],
[393,"15.310.1303","0,50 mm bakır levhadan çatı deresi yapılması ve yerine tesbiti","m",3155.25,3155.25,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",0.9,6.0],["10.100.1026","Tenekeci ustası","Sa",1.2,310.0],["10.100.1064","Çırak","Sa",1.2,205.0]]],
[394,"15.310.1304","0,50 mm bakır levhadan oluk halinde çatı deresi yapılması ve yerine tespiti","m",5487.38,5487.38,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",1.5,6.0],["10.100.1026","Tenekeci ustası","Sa",2.1,310.0],["10.100.1064","Çırak","Sa",2.1,205.0]]],
[395,"15.310.1306","0,50 mm bakır levhadan atika duvarı arkasına çatı deresi yapılması ve yerine konulması","m",5778.75,5778.75,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",1.5,6.0],["10.100.1026","Tenekeci ustası","Sa",2.1,310.0],["10.100.1064","Çırak","Sa",2.1,205.0],["10.330.5495","Bitümlü karton","m²",0.35,6.0]]],
[396,"15.310.1309","0,50 mm bakır levhadan pencere denizliği yapılması ve yerine konulması","m",1980.81,1980.81,"Kaba İnşaat","",[["10.330.5495","Bitümlü karton","m²",0.45,6.0],["10.100.1026","Tenekeci ustası","Sa",1.0,310.0],["10.100.1064","Çırak","Sa",1.0,205.0],["10.420.1401","Sert PVC boru yağmur borusu (conta dahil)","m",1.05,55.0]]],
[397,"15.315.1001","Ø 70 mm çapında bir ucu muflu sert PVC yağmur borusu temini ve yerine tesbiti","m",207.41,207.41,"Kaba İnşaat","",[["10.420.1401","Sert PVC boru yağmur borusu (conta dahil)","m",1.05,55.0],["10.100.1070","İkinci sınıf usta","Sa",0.2,300.0],["10.100.1071","İkinci sınıf usta yardımcısı","Sa",0.1,230.0],["10.420.1402","Sert PVC boru yağmur borusu (conta dahil)","m",1.05,90.0],["10.100.1070","İkinci sınıf usta","Sa",0.2,300.0],["10.100.1071","İkinci sınıf usta yardımcısı","Sa",0.1,230.0]]],
[398,"15.315.1002","Ø 100 mm çapında bir ucu muflu sert PVC yağmur borusu temini ve yerine tesbiti","m",268.14,268.14,"Kaba İnşaat","",[["10.420.1402","Sert PVC boru yağmur borusu (conta dahil)","m",1.05,90.0],["10.100.1070","İkinci sınıf usta","Sa",0.2,300.0],["10.100.1071","İkinci sınıf usta yardımcısı","Sa",0.1,230.0],["10.420.1403","Sert PVC boru yağmur borusu (conta dahil)","m",1.05,110.0],["10.100.1070","İkinci sınıf usta","Sa",0.2,300.0],["10.100.1071","İkinci sınıf usta yardımcısı","Sa",0.1,230.0]]],
[399,"15.315.1003","Ø 125 mm çapında bir ucu muflu sert PVC yağmur borusu temini ve yerine tesbiti","m",310.75,310.75,"Kaba İnşaat","",[["10.420.1403","Sert PVC boru yağmur borusu (conta dahil)","m",1.05,110.0],["10.100.1070","İkinci sınıf usta","Sa",0.2,300.0],["10.100.1071","İkinci sınıf usta yardımcısı","Sa",0.1,230.0],["10.420.1405","Sert PVC yağmur oluğu","m",1.05,38.0],["10.420.1404","Sert PVC çatı bantı","m",1.05,40.0],["10.420.1405","Sert PVC yağmur oluğu","m",0.42,38.0],["10.100.1070","İkinci sınıf usta","Sa",0.3,300.0],["10.100.1071","İkinci sınıf usta yardımcısı","Sa",0.15,230.0]]],
[400,"15.315.1004","Ø 100 mm çapında sert PVC yağmur oluğu temini ve yerine tesbiti","m",371.1,371.1,"Kaba İnşaat","",[["10.420.1405","Sert PVC yağmur oluğu","m",1.05,38.0],["10.420.1404","Sert PVC çatı bantı","m",1.05,40.0],["10.420.1405","Sert PVC yağmur oluğu","m",0.42,38.0],["10.100.1070","İkinci sınıf usta","Sa",0.3,300.0],["10.100.1071","İkinci sınıf usta yardımcısı","Sa",0.15,230.0],["10.420.1406","Sert PVC yağmur oluğu","m",1.05,65.0]]],
[401,"15.315.1005","Ø 150 mm çapında sert PVC yağmur oluğu temini ve yerine tesbiti","m",437.71,437.71,"Kaba İnşaat","",[["10.420.1406","Sert PVC yağmur oluğu","m",1.05,65.0],["10.420.1404","Sert PVC çatı bantı","m",1.05,40.0],["10.420.1406","Sert PVC yağmur oluğu","m",0.42,65.0],["10.100.1070","İkinci sınıf usta","Sa",0.3,300.0],["10.100.1071","İkinci sınıf usta yardımcısı","Sa",0.15,230.0]]],
[402,"15.320.1001","mm kalınlıkta boyalı galvanizli sac ve altı 0.40 mm kalınlıkta boyalı galvanizli sac) çatı","m²",1459.66,1459.66,"Kaba İnşaat","",[["10.330.2602","50 mm poliüretan dolgulu çatı paneli","m²",1.2,750.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3099","EPDM contalı, panel montaj vidası","adet",2.5,3.5],["10.330.3098","Plastik esaslı sızdırmazlık bandı","m",0.1,20.0],["10.380.9982","Silikon (310 ml)","adet",0.1,350.0],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[403,"15.320.1002","mm kalınlıkta PVC membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli ile","m²",2015.91,2015.91,"Kaba İnşaat","",[["10.330.2652","","m²",1.2,1100.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3100","Çektirme vidalı, panel montaj vidası","adet",2.5,3.5],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.2,310.0],["10.100.1010","Yalıtımcı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[404,"15.320.1003","mm kalınlıkta TPO membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli ile","m²",2075.91,2075.91,"Kaba İnşaat","",[["10.330.2677","","m²",1.2,1140.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3100","Çektirme vidalı, panel montaj vidası","adet",2.5,3.5],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.2,310.0],["10.100.1010","Yalıtımcı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[405,"15.320.1004","0.50 mm kalınlıkta boyalı galvanizli sac ve altı 0.40 mm kalınlıkta boyalı galvanizli sac) çatı","m²",1572.16,1572.16,"Kaba İnşaat","",[["10.330.2702","50 mm poliizosiyanurat dolgulu çatı paneli","m²",1.2,825.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3099","EPDM contalı, panel montaj vidası","adet",2.5,3.5],["10.330.3098","Plastik esaslı sızdırmazlık bandı","m",0.1,20.0],["10.380.9982","Silikon (310 ml)","adet",0.1,350.0],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[406,"15.320.1005","1.20 mm kalınlıkta PVC membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli","m²",2165.91,2165.91,"Kaba İnşaat","",[["10.330.2752","","m²",1.2,1200.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3100","Çektirme vidalı, panel montaj vidası","adet",2.5,3.5],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.2,310.0],["10.100.1010","Yalıtımcı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[407,"15.320.1006","1.20 mm kalınlıkta TPO membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli","m²",2180.91,2180.91,"Kaba İnşaat","",[["10.330.2777","","m²",1.2,1210.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3100","Çektirme vidalı, panel montaj vidası","adet",2.5,3.5],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.2,310.0],["10.100.1010","Yalıtımcı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[408,"15.320.1007","mm kalınlıkta, altı 0.50 mm kalınlıkta naturel, gofrajlı alüminyum) çatı paneli ile çatı örtüsü","m²",1684.66,1684.66,"Kaba İnşaat","",[["10.330.2803","60 mm polistren dolgulu çatı paneli","m²",1.2,900.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3099","EPDM contalı, panel montaj vidası","adet",2.5,3.5],["10.330.3098","Plastik esaslı sızdırmazlık bandı","m",0.1,20.0],["10.380.9982","Silikon (310 ml)","adet",0.1,350.0],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[409,"15.320.1008","mm kalınlıkta boyalı galvanizli sac, altı 0.40 mm kalınlıkta boyalı galvanizli sac) çatı paneli","m²",1437.16,1437.16,"Kaba İnşaat","",[["10.330.2828","60 mm polistren dolgulu çatı paneli","m²",1.2,735.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3099","EPDM contalı, panel montaj vidası","adet",2.5,3.5],["10.330.3098","Plastik esaslı sızdırmazlık bandı","m",0.1,20.0],["10.380.9982","Silikon (310 ml)","adet",0.1,350.0],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[410,"15.320.1009","mm kalınlıkta boyalı galvanizli sac, altı 0.40 mm kalınlıkta, naturel, gofrajlı alüminyum) çatı","m²",1557.16,1557.16,"Kaba İnşaat","",[["10.330.2853","60 mm polistren dolgulu çatı paneli","m²",1.2,815.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3099","EPDM contalı, panel montaj vidası","adet",2.5,3.5],["10.330.3098","Plastik esaslı sızdırmazlık bandı","m",0.1,20.0],["10.380.9982","Silikon (310 ml)","adet",0.1,350.0],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[411,"15.320.1010","galvanizli sac ve altı 0.50 mm kalınlıkta boyalı galvanizli sac) çatı paneli ile çatı örtüsü","m²",1854.98,1854.98,"Kaba İnşaat","",[["10.330.2902","60 mm taşyünü dolgulu çatı paneli","m²",1.2,975.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3099","EPDM contalı, panel montaj vidası","adet",2.5,3.5],["10.330.3098","Plastik esaslı sızdırmazlık bandı","m",0.1,20.0],["10.380.9982","Silikon (310 ml)","adet",0.1,350.0],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.25,310.0],["10.100.1062","Düz işçi","Sa",0.75,205.0]]],
[412,"15.320.1014","kalınlıkta PVC membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli ile çatı","m²",2640.6,2640.6,"Kaba İnşaat","",[["10.330.2981","","m²",1.2,1465.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3100","Çektirme vidalı, panel montaj vidası","adet",2.5,3.5],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.25,310.0],["10.100.1010","Yalıtımcı ustası","Sa",0.25,310.0],["10.100.1062","Düz işçi","Sa",0.75,205.0]]],
[413,"15.320.1015","kalınlıkta TPO membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli ile çatı","m²",2693.1,2693.1,"Kaba İnşaat","",[["10.330.2986","","m²",1.2,1500.0],["19.100.1110","Matkap","Sa",0.1,369.82],["10.330.3100","Çektirme vidalı, panel montaj vidası","adet",2.5,3.5],["10.100.1036","Panel Çatı Kaplamacısı","Sa",0.25,310.0],["10.100.1010","Yalıtımcı ustası","Sa",0.25,310.0],["10.100.1062","Düz işçi","Sa",0.75,205.0]]],
[414,"15.325.1001","Ahşap çatı üzerine 0,50 mm kalınlıkta 10 nolu çinkodan çatı örtüsü yapılması","m²",2398.81,2398.81,"Kaba İnşaat","",[["10.330.5493","Bitümlu karton","m²",1.15,5.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.003,9000.0],["10.100.1026","Tenekeci ustası","Sa",2.5,310.0],["10.100.1064","Çırak","Sa",1.5,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[415,"15.325.1002","Ahşap çatı üzerine 0,50 mm kalınlıkta bakır levhadan çatı örtüsü yapılması","m²",5560.13,5560.13,"Kaba İnşaat","",[["10.330.5493","Bitümlu karton","m²",1.15,5.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.003,9000.0],["10.100.1026","Tenekeci ustası","Sa",2.75,310.0],["10.100.1064","Çırak","Sa",1.75,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[416,"15.325.1003","Ahşap çatı üzerine 0,66 mm bakır levhadan çatı örtüsü yapılması","m²",6871.31,6871.31,"Kaba İnşaat","",[["10.330.5493","Bitümlu karton","m²",1.15,5.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.003,9000.0],["10.100.1026","Tenekeci ustası","Sa",2.75,310.0],["10.100.1064","Çırak","Sa",1.75,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[417,"15.325.1007","üzerine 0.70 mm kalınlığında trapezoidal alüminyum levhalar (EN AW 3003 Al-Mn1 Cu) ile","m²",1364.4,1364.4,"Kaba İnşaat","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.006,9000.0],["10.200.2451","Alüminyum trifon takımı","adet",3.0,6.6],["10.200.2452","Alüminyum pop perçin","adet",4.0,0.53],["10.100.1017","Dülger ustası","Sa",0.2,310.0],["10.100.1032","Alüminyum ustası","Sa",0.4,310.0],["10.100.1064","Çırak","Sa",0.4,205.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[418,"15.325.1103","mevcut çatı döşemesi üzerine 0.50 mm kalınlıkta sıcak daldırma galvanizli, oluklu/trapez sac","m²",697.11,697.11,"Kaba İnşaat","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.003,9000.0],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0],["10.240.9116","Kepli trifon vida","adet",2.0,2.7],["10.420.1517","Lastik conta","adet",2.0,2.2],["10.420.1154","Madeni pul","adet",2.0,1.0],["10.100.1017","Dülger ustası","Sa",0.1,310.0],["10.100.1026","Tenekeci ustası","Sa",0.25,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0]]],
[419,"15.325.1104","mevcut çatı döşemesi üzerine elyaflı çimentodan yapılmış oluklu levhalarla çatı örtüsü","m²",631.94,631.94,"Kaba İnşaat","",[["10.240.9101","","m²",1.25,115.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.003,9000.0],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0],["10.240.9116","Kepli trifon vida","adet",2.0,2.7],["10.420.1517","Lastik conta","adet",2.0,2.2],["10.420.1154","Madeni pul","adet",2.0,1.0],["10.100.1017","Dülger ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[420,"15.325.1105","mevcut çatı döşemesi üzerine her renkte oluklu bitümlü levhalarla çatı örtüsü yapılması","m²",554.99,554.99,"Kaba İnşaat","",[["10.240.9102","","m²",1.25,160.0],["10.240.9113","Galvenizli plastik rondelalı burgulu özel çivi","adet",10.0,3.0],["19.100.1110","Matkap","Sa",0.05,369.82],["10.100.1017","Dülger ustası","Sa",0.3,310.0],["10.100.1064","Çırak","Sa",0.25,205.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[421,"15.341.1001","levhalar (EPS ) ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası yapılması","m²",1000.81,1181.25,"İnce İnşaat","",[["10.310.1301","","m³",0.0525,1650.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[422,"15.341.1002","levhalar (EPS) ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası yapılması","m²",1022.46,1222.39,"İnce İnşaat","",[["10.310.1301","","m³",0.063,1650.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[423,"15.341.1003","levhalar (EPS) ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası yapılması","m²",1065.78,1304.65,"İnce İnşaat","",[["10.310.1301","","m³",0.084,1650.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[424,"15.341.1004","levhalar ile (EPS) dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası yapılması","m²",1109.09,1386.91,"İnce İnşaat","",[["10.310.1301","","m³",0.105,1650.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[425,"15.341.1021","ekspande polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası","m²",1020.5,1020.5,"İnce İnşaat","",[["10.310.1311","","m³",0.0525,1950.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[426,"15.341.1022","ekspande polistren (EPS) levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım","m²",1046.09,1046.09,"İnce İnşaat","",[["10.310.1311","","m³",0.063,1950.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[427,"15.341.1023","ekspande polistren levhalar (EPS) ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım","m²",1097.28,1097.28,"İnce İnşaat","",[["10.310.1311","","m³",0.084,1950.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[428,"15.341.1024","ekspande polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası","m²",1148.46,1148.46,"İnce İnşaat","",[["10.310.1311","","m³",0.105,1950.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[429,"15.341.1083","Polistren levhalar (EPS) ile duvarlarda içten ısı yalıtımı ve üzerine ısı yalıtım sıvası","m²",1015.65,1015.65,"İnce İnşaat","",[["10.310.1331","","m³",0.0525,2150.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[430,"15.341.1084","Ekspande Polistren levhalar (EPS) ile duvarlarda içten ısı yalıtımı ve üzerine ısı yalıtım","m²",1068.15,1068.15,"İnce İnşaat","",[["10.310.1332","","m³",0.0525,2950.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[431,"15.341.2001","pürüzlü kanallı ekstrüde polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine","m²",1207.53,1573.88,"İnce İnşaat","",[["10.310.1501","","m³",0.063,4000.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[432,"15.341.2002","pürüzlü kanallı ekstrüde polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı","m²",1345.34,1345.34,"İnce İnşaat","",[["10.310.1502","","m³",0.063,5750.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[433,"15.341.2003","pürüzlü kanallı ekstrüde polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı","m²",1312.53,1312.53,"İnce İnşaat","",[["10.310.1501","","m³",0.084,4000.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[434,"15.341.2004","pürüzlü kanallı ekstrüde polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı","m²",1496.28,1496.28,"İnce İnşaat","",[["10.310.1502","","m³",0.084,5750.0],["10.330.2356","Plastik çivili ısı yalıtım dübeli","adet",6.0,2.0],["10.330.2501","Sıva filesi","m²",1.1,15.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",0.25,9.5],["10.130.9991","Su","m³",0.0025,55.0],["10.100.1010","Yalıtımcı ustası","Sa",1.2,310.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.6,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0]]],
[435,"15.341.2041","kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras çatı","m²",328.6,328.6,"İnce İnşaat","",[["10.310.1541","Basma mukavemeti en az 300 kPa (XPS)","m³",0.0525,4350.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["10.310.1541","Basma mukavemeti en az 300 kPa (XPS)","m³",0.063,4350.0]]],
[436,"15.341.2042","cm kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras","m²",385.69,385.69,"İnce İnşaat","",[["10.310.1541","Basma mukavemeti en az 300 kPa (XPS)","m³",0.063,4350.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[437,"15.341.2043","cm kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras","m²",499.88,499.88,"İnce İnşaat","",[["10.310.1541","Basma mukavemeti en az 300 kPa (XPS)","m³",0.084,4350.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["10.310.1541","Basma mukavemeti en az 300 kPa (XPS)","m³",0.105,4350.0]]],
[438,"15.341.2044","cm kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras","m²",614.06,614.06,"İnce İnşaat","",[["10.310.1541","Basma mukavemeti en az 300 kPa (XPS)","m³",0.105,4350.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[439,"15.341.2045","kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras çatı","m²",322.04,322.04,"İnce İnşaat","",[["10.310.1551","Basma mukavemeti en az 300 kPa (XPS)","m³",0.0525,4250.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["10.310.1551","Basma mukavemeti en az 300 kPa (XPS)","m³",0.063,4250.0]]],
[440,"15.341.2046","kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras çatı","m²",377.81,377.81,"İnce İnşaat","",[["10.310.1551","Basma mukavemeti en az 300 kPa (XPS)","m³",0.063,4250.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[441,"15.341.2047","kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras çatı","m²",489.38,489.38,"İnce İnşaat","",[["10.310.1551","Basma mukavemeti en az 300 kPa (XPS)","m³",0.084,4250.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["10.310.1551","Basma mukavemeti en az 300 kPa (XPS)","m³",0.105,4250.0]]],
[442,"15.341.2048","cm kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras","m²",600.94,600.94,"İnce İnşaat","",[["10.310.1551","Basma mukavemeti en az 300 kPa (XPS)","m³",0.105,4250.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[443,"15.341.2061","cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilmeyen ters teras çatılarda) ısı","m²",600.94,600.94,"İnce İnşaat","",[["10.310.1561","Basma mukavemeti en az 200 kPa (XPS)","m³",0.105,4250.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[444,"15.341.2062","cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilmeyen ters teras çatılarda) ısı","m²",594.38,594.38,"İnce İnşaat","",[["10.310.1571","Basma mukavemeti en az 200 kPa (XPS)","m³",0.105,4200.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[445,"15.341.2063","cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilmeyen ters teras çatılarda) ısı","m²",555.0,555.0,"İnce İnşaat","",[["10.310.1581","Basma mukavemeti en az 200 kPa (XPS)","m³",0.105,3900.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[446,"15.341.2064","cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilen ters teras çatılarda) ısı yalıtımı","m²",646.88,646.88,"İnce İnşaat","",[["10.310.1562","Basma mukavemeti en az 300 kPa (XPS)","m³",0.105,4600.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[447,"15.341.2065","cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilen ters teras çatılarda) ısı yalıtımı","m²",640.31,640.31,"İnce İnşaat","",[["10.310.1572","Basma mukavemeti en az 300 kPa (XPS).","m³",0.105,4550.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[448,"15.341.2066","cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilen ters teras çatılarda) ısı yalıtımı","m²",600.94,600.94,"İnce İnşaat","",[["10.310.1582","Basma mukavemeti en az 300 kPa (XPS).","m³",0.105,4250.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[449,"15.341.2081","basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı","m²",466.25,466.25,"İnce İnşaat","",[["10.310.1591","Basma mukavemeti en az 300 kPa (XPS).","m³",0.063,4500.0],["10.100.1010","Yalıtımcı ustası","Sa",0.25,310.0]]],
[450,"15.341.2082","basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı","m²",454.44,454.44,"İnce İnşaat","",[["10.310.1601","Basma mukavemeti en az 300 kPa (XPS).","m³",0.063,4350.0],["10.100.1010","Yalıtımcı ustası","Sa",0.25,310.0]]],
[451,"15.341.2083","basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı","m²",438.69,438.69,"İnce İnşaat","",[["10.310.1611","Basma mukavemeti en az 300 kPa (XPS)","m³",0.063,4150.0],["10.100.1010","Yalıtımcı ustası","Sa",0.25,310.0]]],
[452,"15.341.2084","basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı","m²",600.13,600.13,"İnce İnşaat","",[["10.310.1592","Basma mukavemeti en az 400 kPa (XPS).","m³",0.063,6200.0],["10.100.1010","Yalıtımcı ustası","Sa",0.25,310.0]]],
[453,"15.341.2085","basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı","m²",572.56,572.56,"İnce İnşaat","",[["10.310.1602","Basma mukavemeti en az 400 kPa (XPS).","m³",0.063,5850.0],["10.100.1010","Yalıtımcı ustası","Sa",0.25,310.0]]],
[454,"15.341.2086","basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı","m²",552.88,552.88,"İnce İnşaat","",[["10.310.1612","Basma mukavemeti en az 400 kPa (XPS)","m³",0.063,5600.0],["10.100.1010","Yalıtımcı ustası","Sa",0.25,310.0]]],
[455,"15.341.3041","yada arakat döşeme betonu üzerinde yüke maruz kalan yüzer döşeme uygulaması vb.) ısı","m²",240.0,240.0,"İnce İnşaat","",[["10.310.1201","2,5 cm kalınlıkta taşyünü levha","m²",1.05,150.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["10.310.1211","2,5 cm kalınlıkta taşyünü levha","m²",1.05,165.0]]],
[456,"15.341.3042","yada arakat döşeme betonu üzerinde yüke maruz kalan yüzer döşeme uygulaması vb.) ısı","m²",259.69,259.69,"İnce İnşaat","",[["10.310.1211","2,5 cm kalınlıkta taşyünü levha","m²",1.05,165.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["10.310.1241","","m³",0.0263,2650.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[457,"15.341.3061","(yüke maruz kalmayan uygulamalarda, yükseltilmiş döşeme uygulaması vb.) ısı yalıtımı","m²",130.25,130.25,"İnce İnşaat","",[["10.310.1241","","m³",0.0263,2650.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[458,"15.341.3062","(yüke maruz kalmayan uygulamalarda, yükseltilmiş döşeme uygulaması vb.) ısı yalıtımı","m²",115.45,115.45,"İnce İnşaat","",[["10.310.1242","","m³",0.0263,2200.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["10.310.1251","","m³",0.063,2000.0]]],
[459,"15.341.3081","döşemelerde ısı ve ses yalıtımı yapılması (Ses yutma katsayısı (ortalama) ≥ 0,80 olan","m²",200.63,200.63,"İnce İnşaat","",[["10.310.1251","","m³",0.063,2000.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[460,"15.341.3082","döşemelerde ısı ve ses yalıtımı yapılması (Ses yutma katsayısı (ortalama) ≥ 0,90 olan","m²",263.63,263.63,"İnce İnşaat","",[["10.310.1252","","m³",0.063,2800.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0],["10.310.1253","","m³",0.063,3500.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[461,"15.341.3083","döşemelerde ısı ve ses yalıtımı yapılması (Ses yutma katsayısı (ortalama) ≥ 1,00 olan","m²",318.75,318.75,"İnce İnşaat","",[["10.310.1253","","m³",0.063,3500.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.15,230.0]]],
[462,"15.341.4001","8 cm kalınlıkta kaplamasız camyünü şilte ve üzerine su buharı geçişine açık su yalıtım örtüsü","m²",289.5,289.5,"İnce İnşaat","",[["10.310.1001","","m³",0.084,1150.0],["10.330.5498","Su buharı geçişine açık su yalıtım örtüsü","m²",1.1,60.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.3,230.0]]],
[463,"15.341.4002","8 cm kalınlıkta kaplamasız camyünü şilte ve üzerine su buharı geçişine açık su yalıtım örtüsü","m²",245.93,245.93,"İnce İnşaat","",[["10.310.1002","","m³",0.084,735.0],["10.330.5498","Su buharı geçişine açık su yalıtım örtüsü","m²",1.1,60.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.3,230.0]]],
[464,"15.341.4003","8 cm kalınlıkta bir yüzü cam tülü kaplamalı camyünü şilte ve üzerine su buharı geçişine açık","m²",326.25,326.25,"İnce İnşaat","",[["10.310.1011","","m²",1.05,120.0],["10.330.5498","Su buharı geçişine açık su yalıtım örtüsü","m²",1.1,60.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.3,230.0]]],
[465,"15.341.4004","8 cm kalınlıkta bir yüzü camtülü kaplamalı camyünü şilte ve üzerine su buharı geçişine açık","m²",273.75,273.75,"İnce İnşaat","",[["10.310.1021","","m²",1.05,80.0],["10.330.5498","Su buharı geçişine açık su yalıtım örtüsü","m²",1.1,60.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.3,230.0]]],
[466,"15.341.4005","16 cm kalınlıkta kaplamasız camyünü şilte ve üzerine su buharı geçişine açık su yalıtım","m²",410.25,410.25,"İnce İnşaat","",[["10.310.1001","","m³",0.168,1150.0],["10.330.5498","Su buharı geçişine açık su yalıtım örtüsü","m²",1.1,60.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.3,230.0]]],
[467,"15.341.4006","16 cm kalınlıkta kaplamasız camyünü şilte ve üzerine su buharı geçişine açık su yalıtım","m²",323.1,323.1,"İnce İnşaat","",[["10.310.1002","","m³",0.168,735.0],["10.330.5498","Su buharı geçişine açık su yalıtım örtüsü","m²",1.1,60.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.3,230.0]]],
[468,"15.341.4007","16 cm kalınlıkta bir yüzü camtülü kaplamalı camyünü şilte ve üzerine su buharı geçişine açık","m²",464.06,464.06,"İnce İnşaat","",[["10.310.1012","","m²",1.05,225.0],["10.330.5498","Su buharı geçişine açık su yalıtım örtüsü","m²",1.1,60.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.3,230.0]]],
[469,"15.341.4008","16 cm kalınlıkta bir yüzü camtülü kaplamalı camyünü şilte ve üzerine su buharı geçişine açık","m²",365.63,365.63,"İnce İnşaat","",[["10.310.1022","","m²",1.05,150.0],["10.330.5498","Su buharı geçişine açık su yalıtım örtüsü","m²",1.1,60.0],["10.100.1042","Yalıtımcı usta yardımcısı","Sa",0.3,230.0]]],
[470,"15.360.1001","Alüminyum köse profilinin (fileli) temini ve yerine tesbiti","m",50.56,50.56,"İnce İnşaat","",[["10.330.2403","Alüminyum Köşe Profilleri (Fileli)","m",1.05,14.0],["10.100.1010","Yalıtımcı ustası","Sa",0.05,310.0],["10.100.1062","Düz işçi","Sa",0.05,205.0],["10.330.2404","PVC Köşe Profilleri (Fileli)","m",1.05,9.5],["10.100.1010","Yalıtımcı ustası","Sa",0.05,310.0],["10.100.1062","Düz işçi","Sa",0.05,205.0]]],
[471,"15.360.1002","PVC köse profilinin (fileli) temini ve yerine tespiti","m",44.66,44.66,"İnce İnşaat","",[["10.330.2404","PVC Köşe Profilleri (Fileli)","m",1.05,9.5],["10.100.1010","Yalıtımcı ustası","Sa",0.05,310.0],["10.100.1062","Düz işçi","Sa",0.05,205.0],["10.330.2407","","m",1.05,28.0]]],
[472,"15.360.1003","Alüminyum damlalıklı köşe profilinin (fileli) temini ve yerine tespiti","m",68.94,68.94,"İnce İnşaat","",[["10.330.2407","","m",1.05,28.0],["10.100.1010","Yalıtımcı ustası","Sa",0.05,310.0],["10.100.1062","Düz işçi","Sa",0.05,205.0],["10.330.2408","PVC Damlalıklı Köşe Profilleri (Fileli)","m",1.05,15.0],["10.100.1010","Yalıtımcı ustası","Sa",0.05,310.0],["10.100.1062","Düz işçi","Sa",0.05,205.0]]],
[473,"15.360.1004","PVC damlalıklı köşe profilinin (fileli) temini ve yerine tespiti","m",51.88,51.88,"İnce İnşaat","",[["10.330.2408","PVC Damlalıklı Köşe Profilleri (Fileli)","m",1.05,15.0],["10.100.1010","Yalıtımcı ustası","Sa",0.05,310.0],["10.100.1062","Düz işçi","Sa",0.05,205.0],["10.330.2411","","m",1.05,45.0]]],
[474,"15.365.1001","tesviyesi yapılması ve üzerine 2 mm kalınlıkta PVC esaslı yer döseme malzemeleri ile döseme","m²",1047.03,1047.03,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6001","PVC esaslı yer döşemesi","m²",1.05,460.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[475,"15.365.1002","tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı yer döseme kaplaması yapılması","m²",948.59,948.59,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6002","PVC esaslı yer döşemesi","m²",1.05,385.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[476,"15.365.1003","tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme","m²",1027.34,1027.34,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6010","PVC esaslı yer döşemesi","m²",1.05,445.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[477,"15.365.1004","tesviyesi yapılması ve üzerine 2 mm kalınlıkta PVC esaslı karo yer döseme malzemeleri ile","m²",1270.15,1270.15,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6012","PVC esaslı yer döşemesi","m²",1.05,605.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[478,"15.365.1005","tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme","m²",1276.71,1276.71,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6021","PVC esaslı yer döşemesi","m²",1.05,635.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[479,"15.365.1006","tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme","m²",1447.34,1447.34,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6022","PVC esaslı yer döşemesi","m²",1.05,765.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[480,"15.365.1007","tesviyesi yapılması ve üzerine 3 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme","m²",1145.46,1145.46,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6031","PVC esaslı yer döşemesi","m²",1.05,535.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[481,"15.365.1008","tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı yer döseme kaplaması yapılması","m²",1191.4,1191.4,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6003","PVC esaslı yer döşemesi","m²",1.05,570.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[482,"15.365.1009","tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı karo yer döseme kaplaması","m²",1394.84,1394.84,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6011","PVC esaslı yer döşemesi","m²",1.05,700.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[483,"15.365.1021","tesviyesi yapılması ve üzerine 2 mm kalınlıkta PVC esaslı yer döseme malzemeleri ile döseme","m²",1025.4,1025.4,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6001","PVC esaslı yer döşemesi","m²",1.05,460.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[484,"15.365.1022","tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı yer döseme kaplaması yapılması","m²",926.96,926.96,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6002","PVC esaslı yer döşemesi","m²",1.05,385.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[485,"15.365.1023","tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme","m²",1005.71,1005.71,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6010","PVC esaslı yer döşemesi","m²",1.05,445.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[486,"15.365.1024","tesviyesi yapılması ve üzerine 2 mm kalınlıkta PVC esaslı karo yer döseme malzemeleri ile","m²",1248.53,1248.53,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6012","PVC esaslı yer döşemesi","m²",1.05,605.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[487,"15.365.1025","tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme","m²",1255.09,1255.09,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6021","PVC esaslı yer döşemesi","m²",1.05,635.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[488,"15.365.1026","tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme","m²",1425.71,1425.71,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6022","PVC esaslı yer döşemesi","m²",1.05,765.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[489,"15.365.1027","tesviyesi yapılması ve üzerine 3 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme","m²",1123.84,1123.84,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6031","PVC esaslı yer döşemesi","m²",1.05,535.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[490,"15.365.1028","tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı yer döseme kaplaması yapılması","m²",1169.78,1169.78,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6003","PVC esaslı yer döşemesi","m²",1.05,570.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[491,"15.365.1029","tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı karo yer döseme kaplaması","m²",1373.21,1373.21,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6011","PVC esaslı yer döşemesi","m²",1.05,700.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[492,"15.365.1101","tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde","m²",2087.65,2087.65,"İnce İnşaat","",[["10.240.6071","PVC esaslı spor zemin kaplaması (P1)","m²",1.05,1150.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.4,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[493,"15.365.1102","tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde","m²",2612.65,2612.65,"İnce İnşaat","",[["10.240.6072","PVC esaslı spor zemin kaplaması (P2)","m²",1.05,1550.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.4,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[494,"15.365.1103","tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde","m²",3268.9,3268.9,"İnce İnşaat","",[["10.240.6073","PVC esaslı spor zemin kaplaması (P3)","m²",1.05,2050.0],["15.190.1007","","m²",1.0,158.17],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.4,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[495,"15.365.1111","tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde","m²",2066.03,2066.03,"İnce İnşaat","",[["10.240.6071","PVC esaslı spor zemin kaplaması (P1)","m²",1.05,1150.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.4,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[496,"15.365.1112","tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde","m²",2591.03,2591.03,"İnce İnşaat","",[["10.240.6072","PVC esaslı spor zemin kaplaması (P2)","m²",1.05,1550.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.4,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[497,"15.365.1113","tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde","m²",3247.28,3247.28,"İnce İnşaat","",[["10.240.6073","PVC esaslı spor zemin kaplaması (P3)","m²",1.05,2050.0],["15.190.1019","","m²",1.0,140.87],["10.240.6053","Kaynak kordonu","m",0.8,9.0],["10.100.1068","Birinci sınıf usta","Sa",0.4,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.4,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[498,"15.365.1501","tesviyesi yapılması ve üzerine 2 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf","m²",1192.09,1192.09,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6101","2 mm kalınlıkta linolyum zemin kaplaması","m²",1.05,570.0],["15.190.1007","","m²",1.0,158.17],["10.240.6104","Linolyum kaynak kordonu","m",0.8,12.5],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[499,"15.365.1502","tesviyesi yapılması ve üzerine 2,5 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf","m²",1119.9,1119.9,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6102","2,5 mm kalınlıkta linolyum zemin kaplaması","m²",1.05,515.0],["15.190.1007","","m²",1.0,158.17],["10.240.6104","Linolyum kaynak kordonu","m",0.8,12.5],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[500,"15.365.1503","tesviyesi yapılması ve üzerine 3,2 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf","m²",1448.03,1448.03,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6103","3,2 mm kalınlıkta linolyum zemin kaplaması","m²",1.05,765.0],["15.190.1007","","m²",1.0,158.17],["10.240.6104","Linolyum kaynak kordonu","m",0.8,12.5],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[501,"15.365.1511","tesviyesi yapılması ve üzerine 2 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf","m²",1170.46,1170.46,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6101","2 mm kalınlıkta linolyum zemin kaplaması","m²",1.05,570.0],["15.190.1019","","m²",1.0,140.87],["10.240.6104","Linolyum kaynak kordonu","m",0.8,12.5],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[502,"15.365.1512","tesviyesi yapılması ve üzerine 2,5 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf","m²",1098.28,1098.28,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6102","2,5 mm kalınlıkta linolyum zemin kaplaması","m²",1.05,515.0],["15.190.1019","","m²",1.0,140.87],["10.240.6104","Linolyum kaynak kordonu","m",0.8,12.5],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[503,"15.365.1513","tesviyesi yapılması ve üzerine 3,2 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf","m²",1426.4,1426.4,"İnce İnşaat","1) Kaplama yapılan yüzeyler projesi üzerinden ölçülür.",[["10.240.6103","3,2 mm kalınlıkta linolyum zemin kaplaması","m²",1.05,765.0],["15.190.1019","","m²",1.0,140.87],["10.240.6104","Linolyum kaynak kordonu","m",0.8,12.5],["10.100.1068","Birinci sınıf usta","Sa",0.2,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",0.2,230.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[504,"15.365.1701","PVC esaslı esnek süpürgelik temini ve yerine tespit edilmesi.","m",72.31,72.31,"İnce İnşaat","",[["10.240.6051","PVC esaslı süpürgelik","m",1.1,38.0],["10.100.1068","Birinci sınıf usta","Sa",0.02,310.0],["10.100.1062","Düz işçi","Sa",0.02,205.0],["10.240.6052","","m",1.1,57.0],["10.100.1068","Birinci sınıf usta","Sa",0.02,310.0]]],
[505,"15.365.1702","PVC esaslı kendinden dönüşlü kepli süpürgelik temini ve yerine tespit edilmesi.","m",98.44,98.44,"İnce İnşaat","",[["10.240.6052","","m",1.1,57.0],["10.100.1068","Birinci sınıf usta","Sa",0.02,310.0],["10.100.1062","Düz işçi","Sa",0.02,205.0]]],
[506,"15.365.1751","PVC esaslı geçiş profili (4 cm genişliğinde) temini ve yerine monte edilmesi.","m",178.88,178.88,"İnce İnşaat","",[["10.240.6054","","m",1.1,51.0],["10.380.9982","Silikon (310 ml)","adet",0.16,350.0],["10.100.1068","Birinci sınıf usta","Sa",0.1,310.0],["10.240.6055","","m",1.1,102.0],["10.380.9982","Silikon (310 ml)","adet",0.16,350.0],["10.100.1068","Birinci sınıf usta","Sa",0.1,310.0]]],
[507,"15.365.1752","Alüminyum esaslı geçiş profili (4 cm genişliğinde) temini ve yerine monte edilmesi.","m",249.0,249.0,"İnce İnşaat","",[["10.240.6055","","m",1.1,102.0],["10.380.9982","Silikon (310 ml)","adet",0.16,350.0],["10.100.1068","Birinci sınıf usta","Sa",0.1,310.0],["10.240.3302","","m²",1.05,171.0]]],
[508,"15.375.1002","I.kalite, beyaz seramik yer karoları ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",628.7,696.39,"İnce İnşaat","",[["10.240.3302","","m²",1.05,171.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[509,"15.375.1004","I.kalite, beyaz seramik yer karoları ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",691.44,691.44,"İnce İnşaat","",[["10.240.3304","","m²",1.06,179.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[510,"15.375.1052","I.kalite, renkli seramik yer karoları ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",639.2,639.2,"İnce İnşaat","",[["10.240.3352","","m²",1.05,179.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[511,"15.375.1054","I.kalite, renkli seramik yer karoları ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",704.69,704.69,"İnce İnşaat","",[["10.240.3354","","m²",1.06,189.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[512,"15.380.1003","I.kalite, beyaz seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması yapılması","m²",714.7,792.5,"İnce İnşaat","",[["10.240.3403","","m²",1.05,207.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[513,"15.380.1005","I.kalite, beyaz seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması yapılması","m²",692.39,692.39,"İnce İnşaat","",[["10.240.3405","","m²",1.05,190.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[514,"15.380.1006","özelliğinde, I.kalite, beyaz seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması","m²",816.31,816.31,"İnce İnşaat","",[["10.240.3406","","m²",1.06,244.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[515,"15.380.1007","türlü desen ve yüzey özelliğinde, I.kalite, beyaz seramik duvar karoları ile 3 mm derz aralıklı","m²",725.33,725.33,"İnce İnşaat","",[["10.240.3407","45 cm anma ebatlarında beyaz seramik duvar","m²",1.05,177.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[516,"15.380.1053","I.kalite, renkli seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması yapılması","m²",734.39,734.39,"İnce İnşaat","",[["10.240.3453","","m²",1.05,222.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[517,"15.380.1055","I.kalite, renkli seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması yapılması","m²",700.26,700.26,"İnce İnşaat","",[["10.240.3455","","m²",1.05,196.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[518,"15.380.1056","özelliğinde, I.kalite, renkli seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması","m²",833.54,833.54,"İnce İnşaat","",[["10.240.3456","","m²",1.06,257.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[519,"15.380.1057","türlü desen ve yüzey özelliğinde, I.kalite, renkli seramik duvar karoları ile 3 mm derz","m²",743.7,743.7,"İnce İnşaat","",[["10.240.3457","45 cm anma ebatlarında renkli seramik duvar","m²",1.05,191.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[520,"15.385.1004","I.kalite, beyaz, sırlı porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",787.83,873.91,"İnce İnşaat","",[["10.240.3504","","m²",1.05,253.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[521,"15.385.1006","I.kalite, beyaz, sırlı porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",809.54,809.54,"İnce İnşaat","",[["10.240.3506","","m²",1.06,267.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[522,"15.385.1024","I.kalite, renkli, sırlı porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",804.89,804.89,"İnce İnşaat","",[["10.240.3554","","m²",1.05,266.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[523,"15.385.1026","I.kalite, renkli, sırlı porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",829.41,829.41,"İnce İnşaat","",[["10.240.3556","","m²",1.06,282.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[524,"15.385.1043","porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı","m²",995.89,995.89,"İnce İnşaat","",[["10.240.3503","","m²",1.05,382.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[525,"15.385.1044","I.kalite, beyaz, sırlı porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması","m²",826.58,826.58,"İnce İnşaat","",[["10.240.3504","","m²",1.05,253.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[526,"15.385.1045","porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı","m²",848.29,848.29,"İnce İnşaat","",[["10.240.3505","","m²",1.06,267.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[527,"15.385.1046","I.kalite, beyaz, sırlı porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması","m²",848.29,848.29,"İnce İnşaat","",[["10.240.3506","","m²",1.06,267.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[528,"15.385.1049","porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı","m²",998.01,998.01,"İnce İnşaat","",[["10.240.3509","","m²",1.06,380.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[529,"15.385.1050","porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı","m²",934.41,934.41,"İnce İnşaat","",[["10.240.3510","","m²",1.06,332.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[530,"15.385.1063","porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı","m²",894.83,894.83,"İnce İnşaat","",[["10.240.3553","","m²",1.05,305.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[531,"15.385.1064","I.kalite, renkli, sırlı porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması","m²",843.64,843.64,"İnce İnşaat","",[["10.240.3554","","m²",1.05,266.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[532,"15.385.1065","porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı","m²",868.16,868.16,"İnce İnşaat","",[["10.240.3555","","m²",1.06,282.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[533,"15.385.1066","I.kalite, renkli, sırlı porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması","m²",868.16,868.16,"İnce İnşaat","",[["10.240.3556","","m²",1.06,282.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[534,"15.385.1069","porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı","m²",1009.94,1009.94,"İnce İnşaat","",[["10.240.3559","","m²",1.06,389.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[535,"15.385.1070","porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı","m²",962.24,962.24,"İnce İnşaat","",[["10.240.3560","","m²",1.06,353.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[536,"15.390.1004","I.kalite, mat, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",833.76,926.16,"İnce İnşaat","",[["10.240.3604","","m²",1.05,288.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[537,"15.390.1006","yüzey özelliğinde, I.kalite, mat, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması","m²",955.29,955.29,"İnce İnşaat","",[["10.240.3606","ebatlarında mat sırsız porselen karo","m²",1.06,377.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[538,"15.390.1008","mat, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",1034.79,1154.81,"İnce İnşaat","",[["10.240.3608","","m²",1.06,437.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[539,"15.390.1009","mat, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",1034.79,1034.79,"İnce İnşaat","",[["10.240.3609","","m²",1.06,437.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[540,"15.390.1010","mat, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",1042.74,1042.74,"İnce İnşaat","",[["10.240.3610","","m²",1.06,443.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[541,"15.390.1024","özelliğinde, I.kalite, parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması","m²",954.51,954.51,"İnce İnşaat","",[["10.240.3654","","m²",1.05,380.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[542,"15.390.1025","parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",997.69,997.69,"İnce İnşaat","",[["10.240.3655","","m²",1.06,409.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[543,"15.390.1026","yüzey özelliğinde, I.kalite, parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme","m²",1111.64,1111.64,"İnce İnşaat","",[["10.240.3656","ebatlarında parlak sırsız porselen karo","m²",1.06,495.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[544,"15.390.1028","parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",1189.81,1189.81,"İnce İnşaat","",[["10.240.3658","","m²",1.06,554.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[545,"15.390.1029","parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",1199.09,1199.09,"İnce İnşaat","",[["10.240.3659","","m²",1.06,561.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[546,"15.390.1030","parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo","m²",1240.16,1240.16,"İnce İnşaat","",[["10.240.3660","","m²",1.06,592.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[547,"15.390.1043","porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı","m²",923.7,923.7,"İnce İnşaat","",[["10.240.3603","","m²",1.05,327.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[548,"15.390.1044","I.kalite, mat, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması","m²",872.51,872.51,"İnce İnşaat","",[["10.240.3604","","m²",1.05,288.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[549,"15.390.1045","porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı","m²",909.24,909.24,"İnce İnşaat","",[["10.240.3605","","m²",1.06,313.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[550,"15.390.1046","yüzey özelliğinde, I.kalite, mat, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe","m²",994.04,994.04,"İnce İnşaat","",[["10.240.3606","ebatlarında mat sırsız porselen karo","m²",1.06,377.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[551,"15.390.1049","mat, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo","m²",1073.54,1073.54,"İnce İnşaat","",[["10.240.3609","","m²",1.06,437.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[552,"15.390.1050","mat, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo","m²",1081.49,1081.49,"İnce İnşaat","",[["10.240.3610","","m²",1.06,443.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[553,"15.390.1063","parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo","m²",1044.45,1044.45,"İnce İnşaat","",[["10.240.3653","","m²",1.05,419.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[554,"15.390.1064","özelliğinde, I.kalite, parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe","m²",993.26,993.26,"İnce İnşaat","",[["10.240.3654","","m²",1.05,380.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[555,"15.390.1065","parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo","m²",1036.44,1036.44,"İnce İnşaat","",[["10.240.3655","","m²",1.06,409.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[556,"15.390.1066","yüzey özelliğinde, I.kalite, parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe","m²",1150.39,1150.39,"İnce İnşaat","",[["10.240.3656","ebatlarında parlak sırsız porselen karo","m²",1.06,495.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[557,"15.390.1069","parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo","m²",1237.84,1237.84,"İnce İnşaat","",[["10.240.3659","","m²",1.06,561.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[558,"15.390.1070","parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo","m²",1278.91,1278.91,"İnce İnşaat","",[["10.240.3660","","m²",1.06,592.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1004","Seramik kaplama ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[559,"15.400.1013","Şartları (Sınıf 2) Yüzey alanı <= 1100cm² ebatlarda ve kırılma dayanımı > 2,5 kN, honlu","m²",1744.75,1744.75,"İnce İnşaat","",[["10.240.4603","Mermer agregalı terrazo karo plak","m²",1.05,354.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[560,"15.400.1014","Şartları (Sınıf 3) 1100 cm²< Yüzey alanı < 1800 cm² ebatlarda ve kırılma dayanımı > 3 kN,","m²",1784.13,1784.13,"İnce İnşaat","",[["10.240.4604","Mermer agregalı terrazo karo plak","m²",1.05,384.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[561,"15.400.1015","Şartları (Sınıf 3) Yüzey alanı >= 1800cm² ebatlarda ve kırılma dayanımı > 3 kN, honlu veya","m²",2007.25,2007.25,"İnce İnşaat","",[["10.240.4605","Mermer agregalı terrazo karo plak","m²",1.05,554.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[562,"15.400.1113","Şartları (Sınıf 2) Yüzey alanı <= 1100cm² ebatlarda ve kırılma dayanımı > 2,5 kN, honlu","m²",1967.88,1967.88,"İnce İnşaat","",[["10.240.4623","Granit agregalı terrazo karo plak","m²",1.05,524.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[563,"15.400.1114","Şartları (Sınıf 3) 1100 cm²< Yüzey alanı < 1800 cm² ebatlarda ve kırılma dayanımı > 3 kN,","m²",2007.25,2007.25,"İnce İnşaat","",[["10.240.4624","Granit agregalı terrazo karo plak","m²",1.05,554.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[564,"15.400.1115","Şartları (Sınıf 3) Yüzey alanı >= 1800cm² ebatlarda ve kırılma dayanımı > 3 kN, honlu veya","m²",2156.88,2156.88,"İnce İnşaat","",[["10.240.4625","Granit agregalı terrazo karo plak","m²",1.05,668.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[565,"15.400.1213","(Kırılma Yükü Şartları (Sınıf 2) Yüzey alanı <= 1100cm² ebatlarda ve kırılma dayanımı >","m²",1967.88,1967.88,"İnce İnşaat","",[["10.240.4643","(min.%20 kuvars/silis + %80 mermer","m²",1.05,524.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[566,"15.400.1214","(Kırılma Yükü Şartları (Sınıf 3) 1100 cm²< Yüzey alanı < 1800 cm² ebatlarda ve kırılma","m²",2007.25,2007.25,"İnce İnşaat","",[["10.240.4644","(min.%20 kuvars/silis + %80 mermer","m²",1.05,554.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[567,"15.400.1215","(Kırılma Yükü Şartları (Sınıf 3) Yüzey alanı >= 1800cm² ebatlarda ve kırılma dayanımı > 3","m²",2156.88,2156.88,"İnce İnşaat","",[["10.240.4645","(min.%20 kuvars/silis + %80 mermer","m²",1.05,668.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[568,"15.400.1313","Şartları (Sınıf 2) Yüzey alanı <= 1100cm² ebatlarda ve kırılma dayanımı > 2,5 kN, honlu","m²",2772.44,2772.44,"İnce İnşaat","",[["10.240.4663","Kuvars/Silis Agregalı terreazo karo plak","m²",1.05,1137.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[569,"15.400.1314","Şartları (Sınıf 3) 1100 cm²< Yüzey alanı < 1800 cm² ebatlarda ve kırılma dayanımı > 3 kN,","m²",2845.94,2845.94,"İnce İnşaat","",[["10.240.4664","Kuvars/Silis Agregalı terreazo karo plak","m²",1.05,1193.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[570,"15.400.1315","Şartları (Sınıf 3) Yüzey alanı >= 1800cm² ebatlarda ve kırılma dayanımı > 3 kN, honlu veya","m²",3015.25,3015.25,"İnce İnşaat","",[["10.240.4665","Kuvars/Silis Agregalı terreazo karo plak","m²",1.05,1322.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[571,"15.405.1011","Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa, Aşınma direnç sınıfı (2-G), Yüzey","m²",1671.25,1671.25,"İnce İnşaat","",[["10.240.4801","Karosiman","m²",1.05,298.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[572,"15.405.1012","Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), 1600 <","m²",1767.06,1767.06,"İnce İnşaat","",[["10.240.4802","Karosiman","m²",1.05,371.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[573,"15.405.1013","Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H), Yüzey alanı","m²",1744.75,1744.75,"İnce İnşaat","",[["10.240.4803","Karosiman","m²",1.05,354.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[574,"15.405.1014","Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),1600 <","m²",1877.31,1877.31,"İnce İnşaat","",[["10.240.4804","Karosiman","m²",1.05,455.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[575,"15.405.1015","Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), Yüzey alanı","m²",1858.94,1858.94,"İnce İnşaat","",[["10.240.4805","Karosiman","m²",1.05,441.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[576,"15.405.1016","Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), 1600 <","m²",1995.44,1995.44,"İnce İnşaat","",[["10.240.4806","Karosiman","m²",1.05,545.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[577,"15.405.1111","Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), Yüzey alanı","m²",1744.75,1744.75,"İnce İnşaat","",[["10.240.4821","Mermer agregalı terrazo karo plak","m²",1.05,354.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[578,"15.405.1112","Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), 1600 <","m²",1840.56,1840.56,"İnce İnşaat","",[["10.240.4822","Mermer agregalı terrazo karo plak","m²",1.05,427.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[579,"15.405.1113","Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H), Yüzey alanı","m²",1840.56,1840.56,"İnce İnşaat","",[["10.240.4823","Mermer agregalı terrazo karo plak","m²",1.05,427.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[580,"15.405.1114","Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),1600 <","m²",1988.88,1988.88,"İnce İnşaat","",[["10.240.4824","Mermer agregalı terrazo karo plak","m²",1.05,540.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[581,"15.405.1115","Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), Yüzey alanı","m²",1969.19,1969.19,"İnce İnşaat","",[["10.240.4825","Mermer agregalı terrazo karo plak","m²",1.05,525.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[582,"15.405.1116","Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), 1600 <","m²",2083.38,2083.38,"İnce İnşaat","",[["10.240.4826","Mermer agregalı terrazo karo plak","m²",1.05,612.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[583,"15.405.1211","Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), Yüzey alanı","m²",1877.31,1877.31,"İnce İnşaat","",[["10.240.4841","Granit agregalı terrazo karo plak","m²",1.05,455.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[584,"15.405.1212","Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), 1600 <","m²",1988.88,1988.88,"İnce İnşaat","",[["10.240.4842","Granit agregalı terrazo karo plak","m²",1.05,540.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[585,"15.405.1213","Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H), Yüzey alanı","m²",1988.88,1988.88,"İnce İnşaat","",[["10.240.4843","Granit agregalı terrazo karo plak","m²",1.05,540.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[586,"15.405.1214","Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),1600 <","m²",2117.5,2117.5,"İnce İnşaat","",[["10.240.4844","Granit agregalı terrazo karo plak","m²",1.05,638.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[587,"15.405.1215","Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), Yüzey alanı","m²",2083.38,2083.38,"İnce İnşaat","",[["10.240.4845","Granit agregalı terrazo karo plak","m²",1.05,612.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[588,"15.405.1216","Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), 1600 <","m²",2231.69,2231.69,"İnce İnşaat","",[["10.240.4846","Granit agregalı terrazo karo plak","m²",1.05,725.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[589,"15.405.1311","Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), Yüzey alanı","m²",1840.56,1840.56,"İnce İnşaat","",[["10.240.4861","Andezit agregalı terrazo karo plak","m²",1.05,427.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[590,"15.405.1312","Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), 1600 <","m²",1969.19,1969.19,"İnce İnşaat","",[["10.240.4862","Andezit agregalı terrazo karo plak","m²",1.05,525.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[591,"15.405.1313","Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H), Yüzey alanı","m²",1895.69,1895.69,"İnce İnşaat","",[["10.240.4863","Andezit agregalı terrazo karo plak","m²",1.05,469.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[592,"15.405.1314","Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),1600 <","m²",2044.0,2044.0,"İnce İnşaat","",[["10.240.4864","Andezit agregalı terrazo karo plak","m²",1.05,582.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[593,"15.405.1315","Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), Yüzey alanı","m²",2019.06,2019.06,"İnce İnşaat","",[["10.240.4865","Andezit agregalı terrazo karo plak","m²",1.05,563.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[594,"15.405.1316","Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), 1600 <","m²",2158.19,2158.19,"İnce İnşaat","",[["10.240.4866","Andezit agregalı terrazo karo plak","m²",1.05,669.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[595,"15.405.1411","Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), Yüzey alanı","m²",1784.13,1784.13,"İnce İnşaat","",[["10.240.4881","Bazalt agregalı terrazo karo plak","m²",1.05,384.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[596,"15.405.1412","Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), 1600 <","m²",1877.31,1877.31,"İnce İnşaat","",[["10.240.4882","Bazalt agregalı terrazo karo plak","m²",1.05,455.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[597,"15.405.1413","Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H), Yüzey alanı","m²",1895.69,1895.69,"İnce İnşaat","",[["10.240.4883","Bazalt agregalı terrazo karo plak","m²",1.05,469.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[598,"15.405.1414","Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),1600 cm² <","m²",2044.0,2044.0,"İnce İnşaat","",[["10.240.4884","Bazalt agregalı terrazo karo plak","m²",1.05,582.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[599,"15.405.1415","Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), Yüzey alanı","m²",1995.44,1995.44,"İnce İnşaat","",[["10.240.4885","Bazalt agregalı terrazo karo plak","m²",1.05,545.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[600,"15.405.1416","Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), 1600 <","m²",2135.88,2135.88,"İnce İnşaat","",[["10.240.4886","Bazalt agregalı terrazo karo plak","m²",1.05,652.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[601,"15.405.1511","Dayanımı Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G),","m²",2065.0,2065.0,"İnce İnşaat","",[["10.240.4901","Kuvars-silis agregalı terrazo karo plak","m²",1.05,598.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[602,"15.405.1512","Dayanımı Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G),","m²",2175.25,2175.25,"İnce İnşaat","",[["10.240.4902","Kuvars-silis agregalı terrazo karo plak","m²",1.05,682.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[603,"15.405.1513","Dayanımı Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),","m²",2175.25,2175.25,"İnce İnşaat","",[["10.240.4903","Kuvars-silis agregalı terrazo karo plak","m²",1.05,682.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[604,"15.405.1514","Dayanımı Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı","m²",2250.06,2250.06,"İnce İnşaat","",[["10.240.4904","Kuvars-silis agregalı terrazo karo plak","m²",1.05,739.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[605,"15.405.1515","Dayanımı Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I),","m²",2250.06,2250.06,"İnce İnşaat","",[["10.240.4905","Kuvars-silis agregalı terrazo karo plak","m²",1.05,739.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[606,"15.405.1516","Dayanımı Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I),","m²",2341.94,2341.94,"İnce İnşaat","",[["10.240.4906","Kuvars-silis agregalı terrazo karo plak","m²",1.05,809.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[607,"15.405.1611","Yükü Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G),","m²",1729.0,1729.0,"İnce İnşaat","",[["10.240.4921","Washbeton terrazo karo plak","m²",1.05,342.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[608,"15.405.1612","Yükü Sartları (Sınıf 1) Egilme dayanımı minimum 2,8 Mpa Asınma direnç sınıfı (2-G), 1600","m²",1840.56,1840.56,"İnce İnşaat","",[["10.240.4922","Washbeton terrazo karo plak","m²",1.05,427.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[609,"15.405.1613","Dayanımı Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),","m²",1969.19,1969.19,"İnce İnşaat","",[["10.240.4923","Washbeton terrazo karo plak","m²",1.05,525.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[610,"15.405.1614","Dayanımı Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı","m²",2025.63,2025.63,"İnce İnşaat","",[["10.240.4924","Washbeton terrazo karo plak","m²",1.05,568.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[611,"15.405.1615","Dayanımı Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I),","m²",2007.25,2007.25,"İnce İnşaat","",[["10.240.4925","Washbeton terrazo karo plak","m²",1.05,554.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[612,"15.405.1616","Dayanımı Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I),","m²",2117.5,2117.5,"İnce İnşaat","",[["10.240.4926","Washbeton terrazo karo plak","m²",1.05,638.0],["19.100.2510","","m³",0.04,2151.3],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",1.6,310.0],["10.100.1062","Düz işçi","Sa",1.6,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[613,"15.420.1011","4 cm kalınlığında andezit levha ile döşeme kaplaması yapılması (30cmxserbest boy)","m²",2414.94,2414.94,"İnce İnşaat","",[["10.240.2668","Andezit plak (4x30xserbest boy cm)","m²",1.05,677.0],["19.100.2510","","m³",0.04,2151.3],["19.100.2517","600 dozlu çimento şerbeti","m³",0.001,1995.0],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",2.0,310.0],["10.100.1062","Düz işçi","Sa",2.0,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[614,"15.420.1111","3 cm kalınlığında andezit levha ile duvar kaplaması yapılması (30cmxserbest boy)","m²",2584.28,2584.28,"İnce İnşaat","1)Andezit kaplanan bütün yüzeyler projesi üzerinden hesaplanır.",[["10.240.2648","Andezit plak","m²",1.05,566.0],["19.100.2510","","m³",0.025,2151.3],["19.100.2518","500 dozlu çimento şerbeti","m³",0.005,1777.6],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",2.5,310.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[615,"15.420.1211","3 cm kalınlığında andezit levha ile söve yapılması","m²",2621.68,2621.68,"İnce İnşaat","",[["10.240.2648","Andezit plak","m²",1.1,566.0],["19.100.2510","","m³",0.035,2151.3],["19.100.2518","500 dozlu çimento şerbeti","m³",0.005,1777.6],["10.130.9991","Su","m³",0.01,55.0],["10.100.1005","Mermer kaplama ustası","Sa",2.5,310.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[616,"15.435.1215","10x15x50 cm boyutlarında andezit bordür temini ve yerine döşenmesi","m",767.41,767.41,"İnce İnşaat","",[["10.240.2602","Andezit bordür","adet",2.1,230.0],["19.100.2511","400 dozlu ince harç","m³",0.001,2182.8],["10.100.1013","Duvarcı ustası","Sa",0.25,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.240.2603","10x20x50 cm, Andezit bordür (TS 10835)","adet",2.1,254.0],["19.100.2511","400 dozlu ince harç","m³",0.001,2182.8]]],
[617,"15.435.1216","10x20x50 cm boyutlarında andezit bordür temini ve yerine döşenmesi","m",830.41,830.41,"İnce İnşaat","",[["10.240.2603","10x20x50 cm, Andezit bordür (TS 10835)","adet",2.1,254.0],["19.100.2511","400 dozlu ince harç","m³",0.001,2182.8],["10.100.1013","Duvarcı ustası","Sa",0.25,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.240.2613","Andezit bordür","adet",1.5,360.0],["19.100.2511","400 dozlu ince harç","m³",0.001,2182.8],["10.100.1013","Duvarcı ustası","Sa",0.25,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[618,"15.435.1217","10x20x70 cm boyutlarında andezit bordür temini ve yerine döşenmesi","m",838.66,838.66,"İnce İnşaat","",[["10.240.2613","Andezit bordür","adet",1.5,360.0],["19.100.2511","400 dozlu ince harç","m³",0.001,2182.8],["10.100.1013","Duvarcı ustası","Sa",0.25,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.480.1101","","m",1.05,185.0]]],
[619,"15.435.1313","50 x 20 cm boyutlarında andezit oluk tası döşenmesi","m",957.73,957.73,"İnce İnşaat","",[["10.240.2621","8x20x50 cm Andezit oluk taşı (TS 10835)","m",1.05,605.0],["19.100.2511","400 dozlu ince harç","m³",0.001,2182.8],["10.100.1013","Duvarcı ustası","Sa",0.25,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.480.1202","Andezit parke taşı (10*10*10 cm)","ton",0.18,1300.0],["10.130.1004","Kum (elenmesi gerekmeyen ince agrega)","m³",0.15,137.0],["10.100.1014","Kaldırımcı ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.9,205.0]]],
[620,"15.440.1001","duvar, tavan ve cephelerde kaplama üstü dilatasyon fugası yapılması (50 mm genislikte","m",403.41,403.41,"İnce İnşaat","",[["10.200.2701","","m",1.05,151.0],["10.420.1012","Vida ve plastik dübel","adet",3.0,2.0],["10.200.2791","Butil bant","m",1.0,18.2],["19.100.1110","Matkap","Sa",0.1,369.82],["10.100.1032","Alüminyum ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[621,"15.440.1002","mm hareket kapasiteli, profil yüksekligi min. 13mm, kanat genisligi min.45 mm) duvar ve","m",588.2,588.2,"İnce İnşaat","",[["10.200.2741","","m",1.05,172.0],["10.420.1012","Vida ve plastik dübel","adet",5.0,2.0],["19.100.1110","Matkap","Sa",0.2,369.82],["10.100.1032","Alüminyum ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[622,"15.440.1003","zeminlerde kaplama üstü dilatasyon fugası yapılması (50 mm genislikte dilatasyonlar için)","m",493.98,493.98,"İnce İnşaat","",[["10.200.2711","","m",1.05,220.0],["10.420.1012","Vida ve plastik dübel","adet",3.0,2.0],["10.200.2791","Butil bant","m",1.0,18.2],["19.100.1110","Matkap","Sa",0.1,369.82],["10.100.1032","Alüminyum ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[623,"15.440.1004","hareket kapasiteli, profil yüksekligi min. 35 mm, kanat genisligi min. 45mm) zeminde","m",873.7,873.7,"İnce İnşaat","",[["10.200.2721","","m",1.05,321.0],["10.200.4024","","adet",7.0,4.35],["19.100.1110","Matkap","Sa",0.2,369.82],["10.100.1032","Alüminyum ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[624,"15.440.1005","genisligi min. 45mm, fitilin yerlestirildigi mesnetler arası ilave elemanlarla güçlendirilmis)","m",1304.2,1304.2,"İnce İnşaat","",[["10.200.2731","","m",1.05,649.0],["10.200.4024","","adet",7.0,4.35],["19.100.1110","Matkap","Sa",0.2,369.82],["10.100.1032","Alüminyum ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[625,"15.440.1006","mm hareket kapasiteli, profil yüksekligi min. 15 mm, kanat genisligi min. 45mm) zeminde","m",816.39,816.39,"İnce İnşaat","",[["10.200.2751","","m",1.05,293.0],["10.420.1012","Vida ve plastik dübel","adet",7.0,2.0],["19.100.1110","Matkap","Sa",0.2,369.82],["10.100.1032","Alüminyum ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[626,"15.440.1007","kanat genisligi min. 45mm, fitilin yerlestirildigi mesnetler arası ilave elemanlarla","m",913.51,913.51,"İnce İnşaat","",[["10.200.2761","","m",1.05,367.0],["10.420.1012","Vida ve plastik dübel","adet",7.0,2.0],["19.100.1110","Matkap","Sa",0.2,369.82],["10.100.1032","Alüminyum ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[627,"15.450.1011","Mozayik denizlik yapılması (normal çimentolu)","m²",5593.55,5593.55,"İnce İnşaat","",[["19.100.2501","200 Kg çimento dozlu harç yapılması","m³",0.03,1779.65],["19.100.2525","Mozayik harcı","m³",0.03,3205.5],["10.130.9991","Su","m³",0.005,55.0],["10.100.1007","Mozayikçi ustası","Sa",8.0,310.0],["10.100.1062","Düz işçi","Sa",8.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["19.100.2501","200 Kg çimento dozlu harç yapılması","m³",0.03,1779.65],["19.100.2426","Mozayik harcı","m³",0.03,4863.0]]],
[628,"15.450.1012","Mozayik denizlik yapılması (beyaz çimentolu)","m²",5655.7,5655.7,"İnce İnşaat","",[["19.100.2501","200 Kg çimento dozlu harç yapılması","m³",0.03,1779.65],["19.100.2426","Mozayik harcı","m³",0.03,4863.0],["10.130.9991","Su","m³",0.005,55.0],["10.100.1007","Mozayikçi ustası","Sa",8.0,310.0],["10.100.1062","Düz işçi","Sa",8.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[629,"15.450.1013","Mozayik parapet yapılması (normal çimentolu)","m²",5562.39,5562.39,"İnce İnşaat","",[["19.100.2501","200 Kg çimento dozlu harç yapılması","m³",0.025,1779.65],["19.100.2525","Mozayik harcı","m³",0.025,3205.5],["10.130.9991","Su","m³",0.005,55.0],["10.100.1007","Mozayikçi ustası","Sa",8.0,310.0],["10.100.1062","Düz işçi","Sa",8.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["19.100.2501","200 Kg çimento dozlu harç yapılması","m³",0.025,1779.65],["19.100.2426","Mozayik harcı","m³",0.025,4863.0]]],
[630,"15.450.1014","Mozaik parapet yapılması (beyaz çimentolu)","m²",5614.19,5614.19,"İnce İnşaat","",[["19.100.2501","200 Kg çimento dozlu harç yapılması","m³",0.025,1779.65],["19.100.2426","Mozayik harcı","m³",0.025,4863.0],["10.130.9991","Su","m³",0.005,55.0],["10.100.1007","Mozayikçi ustası","Sa",8.0,310.0],["10.100.1062","Düz işçi","Sa",8.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[631,"15.465.1001","Gömme iç kapı kilidinin yerine takılması","adet",137.5,145.05,"İnce İnşaat","",[["10.400.2001","Gömme iç kapı kilidi (Geniş tip)","adet",1.0,110.0],["10.400.2002","Gömme iç kapı kilidi (Dar tip)","adet",1.0,110.0],["10.400.2003","","adet",1.0,200.0]]],
[632,"15.465.1002","Gömme iç kapı kilidinin yerine takılması","adet",137.5,137.5,"İnce İnşaat","",[["10.400.2002","Gömme iç kapı kilidi (Dar tip)","adet",1.0,110.0],["10.400.2003","","adet",1.0,200.0],["10.400.2004","","adet",1.0,300.0]]],
[633,"15.465.1003","Gömme makaralı iç kapı kilidinin yerine takılması","adet",250.0,250.0,"İnce İnşaat","",[["10.400.2003","","adet",1.0,200.0],["10.400.2004","","adet",1.0,300.0],["10.400.2005","","adet",1.0,300.0]]],
[634,"15.465.1004","Gömme silindirli iç ve dış kapı kilidinin yerine takılması","adet",375.0,375.0,"İnce İnşaat","",[["10.400.2004","","adet",1.0,300.0],["10.400.2005","","adet",1.0,300.0],["10.400.2006","","adet",1.0,300.0]]],
[635,"15.465.1005","Gömme makaralı silindirli iç ve dış kapı kilidinin yerine takılması","adet",375.0,375.0,"İnce İnşaat","",[["10.400.2005","","adet",1.0,300.0],["10.400.2006","","adet",1.0,300.0],["10.400.2007","Silindir tirajlı dış kapı kilidi","adet",1.0,350.0]]],
[636,"15.465.1006","Gömme makaralı silindirli iç ve dış kapı kilidinin yerine takılması","adet",375.0,375.0,"İnce İnşaat","",[["10.400.2006","","adet",1.0,300.0],["10.400.2007","Silindir tirajlı dış kapı kilidi","adet",1.0,350.0],["10.400.2008","Kapı kolu ve aynaları (Kromajlı)","adet",1.0,110.0],["10.400.2009","Lastik başlı tampon","adet",1.0,18.5]]],
[637,"15.465.1007","Silindirli traşlı dış kapı kilidinin yerine takılması","adet",437.5,437.5,"İnce İnşaat","",[["10.400.2007","Silindir tirajlı dış kapı kilidi","adet",1.0,350.0],["10.400.2008","Kapı kolu ve aynaları (Kromajlı)","adet",1.0,110.0],["10.400.2009","Lastik başlı tampon","adet",1.0,18.5]]],
[638,"15.465.1008","Kapı kolu ve aynalarının yerine takılması","adet",137.5,137.5,"İnce İnşaat","",[["10.400.2008","Kapı kolu ve aynaları (Kromajlı)","adet",1.0,110.0],["10.400.2009","Lastik başlı tampon","adet",1.0,18.5],["10.400.2010","Menteşe","adet",1.0,18.5]]],
[639,"15.465.1009","Lastik başlı tamponun yerine takılması","adet",23.13,23.13,"İnce İnşaat","",[["10.400.2009","Lastik başlı tampon","adet",1.0,18.5],["10.400.2010","Menteşe","adet",1.0,18.5],["10.400.2011","Yaylı menteşe","adet",1.0,170.0]]],
[640,"15.465.1010","Menteşenin yerine takılması","adet",23.13,23.13,"İnce İnşaat","",[["10.400.2010","Menteşe","adet",1.0,18.5],["10.400.2011","Yaylı menteşe","adet",1.0,170.0],["10.400.2012","Sürgü (Düşey tesbit takımı)","adet",1.0,25.0],["10.400.2013","Stop (Nikelajlı)","adet",1.0,90.0]]],
[641,"15.465.1011","Yaylı menteşenin yerine takılması","adet",212.5,212.5,"İnce İnşaat","",[["10.400.2011","Yaylı menteşe","adet",1.0,170.0],["10.400.2012","Sürgü (Düşey tesbit takımı)","adet",1.0,25.0],["10.400.2013","Stop (Nikelajlı)","adet",1.0,90.0],["10.400.2101","İspanyolet takımı (Kol demir ve teferruatı)","adet",1.0,100.0]]],
[642,"15.465.1012","Sürgünün yerine takılması","adet",31.25,31.25,"İnce İnşaat","",[["10.400.2012","Sürgü (Düşey tesbit takımı)","adet",1.0,25.0],["10.400.2013","Stop (Nikelajlı)","adet",1.0,90.0],["10.400.2101","İspanyolet takımı (Kol demir ve teferruatı)","adet",1.0,100.0]]],
[643,"15.465.1013","Stopun yerine takılması","adet",112.5,112.5,"İnce İnşaat","",[["10.400.2013","Stop (Nikelajlı)","adet",1.0,90.0],["10.400.2101","İspanyolet takımı (Kol demir ve teferruatı)","adet",1.0,100.0],["10.400.2102","Vasistas takımı (Basit makas)","adet",1.0,28.0]]],
[644,"15.465.1101","İspanyolet takımının yerine takılması","adet",125.0,125.0,"İnce İnşaat","",[["10.400.2101","İspanyolet takımı (Kol demir ve teferruatı)","adet",1.0,100.0],["10.400.2102","Vasistas takımı (Basit makas)","adet",1.0,28.0],["10.400.2103","","adet",1.0,75.0]]],
[645,"15.465.1102","Vasistas takımının yerine takılması","adet",35.0,35.0,"İnce İnşaat","",[["10.400.2102","Vasistas takımı (Basit makas)","adet",1.0,28.0],["10.400.2103","","adet",1.0,75.0],["10.400.2104","","adet",1.0,60.0]]],
[646,"15.465.1103","Vasistas takımının yerine takılması","adet",93.75,93.75,"İnce İnşaat","",[["10.400.2103","","adet",1.0,75.0],["10.400.2104","","adet",1.0,60.0],["10.400.2105","Sürgü","adet",1.0,20.0]]],
[647,"15.465.1104","Mandalın yerine takılması","adet",75.0,75.0,"İnce İnşaat","",[["10.400.2104","","adet",1.0,60.0],["10.400.2105","Sürgü","adet",1.0,20.0],["10.400.2106","Lastik başlı tampon","adet",1.0,25.0]]],
[648,"15.465.1105","Sürgünün yerine takılması","adet",25.0,25.0,"İnce İnşaat","",[["10.400.2105","Sürgü","adet",1.0,20.0],["10.400.2106","Lastik başlı tampon","adet",1.0,25.0],["10.400.2107","Tespit yaylı mandalı","adet",1.0,30.0]]],
[649,"15.465.1106","Lastik başlı tamponun yerine takılması","adet",31.25,31.25,"İnce İnşaat","",[["10.400.2106","Lastik başlı tampon","adet",1.0,25.0],["10.400.2107","Tespit yaylı mandalı","adet",1.0,30.0]]],
[650,"15.465.1107","Yaylı tespit mandalının yerine takılması","adet",37.5,37.5,"İnce İnşaat","",[["10.400.2107","Tespit yaylı mandalı","adet",1.0,30.0],["10.400.2109","Sürme kanat tutamağı","adet",1.0,85.0],["10.400.2121","80 cm lik 2 kavramalı ispanyolet takımı","adet",1.0,90.0]]],
[651,"15.465.1109","Sürme kanat tutamağının yerine takılması","adet",106.25,106.25,"İnce İnşaat","",[["10.400.2109","Sürme kanat tutamağı","adet",1.0,85.0],["10.400.2121","80 cm lik 2 kavramalı ispanyolet takımı","adet",1.0,90.0],["10.400.2122","100 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,100.0]]],
[652,"15.465.1110","Kavramalı ispanyolet takımının yerine takılması","adet",112.5,112.5,"İnce İnşaat","",[["10.400.2121","80 cm lik 2 kavramalı ispanyolet takımı","adet",1.0,90.0],["10.400.2122","100 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,100.0],["10.400.2123","120 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,120.0]]],
[653,"15.465.1111","Kavramalı ispanyolet takımının yerine takılması","adet",125.0,125.0,"İnce İnşaat","",[["10.400.2122","100 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,100.0],["10.400.2123","120 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,120.0],["10.400.2124","140 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,120.0]]],
[654,"15.465.1112","Kavramalı ispanyolet takımının yerine takılması","adet",150.0,150.0,"İnce İnşaat","",[["10.400.2123","120 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,120.0],["10.400.2124","140 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,120.0],["10.400.2125","160 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,130.0],["10.400.2126","180 cm lik 4 kavramalı ispanyolet takımı","adet",1.0,140.0]]],
[655,"15.465.1113","Kavramalı ispanyolet takımının yerine takılması","adet",150.0,150.0,"İnce İnşaat","",[["10.400.2124","140 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,120.0],["10.400.2125","160 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,130.0],["10.400.2126","180 cm lik 4 kavramalı ispanyolet takımı","adet",1.0,140.0],["10.400.2127","Menteşe","adet",1.0,25.0]]],
[656,"15.465.1114","Kavramalı ispanyolet takımının yerine takılması","adet",162.5,162.5,"İnce İnşaat","",[["10.400.2125","160 cm lik 3 kavramalı ispanyolet takımı","adet",1.0,130.0],["10.400.2126","180 cm lik 4 kavramalı ispanyolet takımı","adet",1.0,140.0],["10.400.2127","Menteşe","adet",1.0,25.0]]],
[657,"15.465.1115","Kavramalı ispanyolet takımının yerine takılması","adet",175.0,175.0,"İnce İnşaat","",[["10.400.2126","180 cm lik 4 kavramalı ispanyolet takımı","adet",1.0,140.0],["10.400.2127","Menteşe","adet",1.0,25.0],["10.400.2128","Boy menteşe","m",1.0,40.0]]],
[658,"15.465.1116","Menteşenin yerine takılması","adet",31.25,31.25,"İnce İnşaat","",[["10.400.2127","Menteşe","adet",1.0,25.0],["10.400.2128","Boy menteşe","m",1.0,40.0],["10.400.2129","Ayarlı menteşe (Çift)","adet",1.0,90.0]]],
[659,"15.465.1117","Boy menteşenin yerine takılması","m",50.0,50.0,"İnce İnşaat","",[["10.400.2128","Boy menteşe","m",1.0,40.0],["10.400.2129","Ayarlı menteşe (Çift)","adet",1.0,90.0],["10.400.2141","","adet",1.0,350.0]]],
[660,"15.465.1118","Ayarlı menteşe (çift) plastik kaplamalının yerine takılması","adet",112.5,112.5,"İnce İnşaat","",[["10.400.2129","Ayarlı menteşe (Çift)","adet",1.0,90.0],["10.400.2141","","adet",1.0,350.0],["10.400.2142","","adet",1.0,400.0]]],
[661,"15.465.1201","İspanyolet takımının yerine takılması (kol dahil) 100 cm'e kadar, 2 kavramalı","adet",437.5,437.5,"İnce İnşaat","",[["10.400.2141","","adet",1.0,350.0],["10.400.2142","","adet",1.0,400.0],["10.400.2143","","adet",1.0,400.0]]],
[662,"15.465.1202","İspanyolet takımının yerine takılması (kol dahil) 180 cm'e kadar, 3 kavramalı","adet",500.0,500.0,"İnce İnşaat","",[["10.400.2142","","adet",1.0,400.0],["10.400.2143","","adet",1.0,400.0],["10.400.2144","Vasistas ispanyolet takımı (kol makas dahil)","adet",1.0,350.0]]],
[663,"15.465.1203","İspanyolet takımının yerine takılması (kol dahil) 180 cm'den büyük, 4 kavramalı","adet",500.0,500.0,"İnce İnşaat","",[["10.400.2143","","adet",1.0,400.0],["10.400.2144","Vasistas ispanyolet takımı (kol makas dahil)","adet",1.0,350.0],["10.380.9981","Camlama takozu","adet",12.0,4.0],["10.380.1511","3+3 mm kal., 12 mm boşluklu yalıtım camı","m²",1.05,955.0],["10.420.1151","Prinç ağac vidası","adet",16.0,0.5],["10.380.9983","Slikon (310 ml) (Asitsiz - Nötral Silikon)","adet",0.8,600.0],["10.100.1022","Camcı ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[664,"15.465.1204","Vasistas ispanyolet takımının yerine takılması (Kol, makas dahil)","adet",437.5,437.5,"İnce İnşaat","",[["10.400.2144","Vasistas ispanyolet takımı (kol makas dahil)","adet",1.0,350.0],["10.380.9981","Camlama takozu","adet",12.0,4.0],["10.380.1511","3+3 mm kal., 12 mm boşluklu yalıtım camı","m²",1.05,955.0],["10.420.1151","Prinç ağac vidası","adet",16.0,0.5],["10.380.9983","Slikon (310 ml) (Asitsiz - Nötral Silikon)","adet",0.8,600.0],["10.100.1022","Camcı ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[665,"15.475.1001","Kadronlu ahşap döşeme yapılması","m²",1169.3,1169.3,"İnce İnşaat","",[["10.130.4501","Çam kerestesi (I.sınıf)","m³",0.03,10200.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.01,9000.0],["19.100.1091","","Sa",0.02,6059.7],["10.100.1017","Dülger ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.8,205.0],["10.130.4501","Çam kerestesi (I.sınıf)","m³",0.03,10200.0]]],
[666,"15.475.1002","Mevcut kadronlar üzerine ahşap döşeme yapılması","m²",953.0,953.0,"İnce İnşaat","",[["10.130.4501","Çam kerestesi (I.sınıf)","m³",0.03,10200.0],["19.100.1091","","Sa",0.015,6059.7],["10.100.1017","Dülger ustası","Sa",0.7,310.0],["10.100.1062","Düz işçi","Sa",0.7,205.0],["10.170.1001","Meşe parke (I.sınıf)","m²",1.05,983.0]]],
[667,"15.480.1001","Beton zemin üzerine kadronlu 15-16 mm kalınlıkta I.sınıf meşe parke kaplama yapılması","m²",2865.46,3106.85,"İnce İnşaat","",[["10.170.1001","Meşe parke (I.sınıf)","m²",1.05,983.0],["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.04,9000.0],["10.420.1852","Hafif agrega","m³",0.05,5.5],["19.100.1091","","Sa",0.02,6059.7],["10.100.1017","Dülger ustası","Sa",1.5,310.0],["10.100.1062","Düz işçi","Sa",1.5,205.0]]],
[668,"15.485.1001","Lamine parke döşeme kaplaması yapılması (süpürgelik dahil)","m²",1712.88,1712.88,"İnce İnşaat","",[["10.170.1402","Lamine parke","m²",1.1,1090.0],["10.330.3501","2 mm kalınlıkta parke altı şiltesi","m²",1.05,6.0],["10.100.1009","Marangoz ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.170.1201","AC1 Sınıf 21 Laminat parke","m²",1.1,270.0],["10.330.3501","2 mm kalınlıkta parke altı şiltesi","m²",1.05,6.0],["10.100.1009","Marangoz ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.15,205.0]]],
[669,"15.490.1001","Laminat parke döşeme kaplaması yapılması (AC1 Sınıf 21) (süpürgelik dahil)","m²",533.81,584.33,"İnce İnşaat","",[["10.170.1201","AC1 Sınıf 21 Laminat parke","m²",1.1,270.0],["10.330.3501","2 mm kalınlıkta parke altı şiltesi","m²",1.05,6.0],["10.100.1009","Marangoz ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.15,205.0],["10.170.1202","AC3 Sınıf 23-31 Laminat parke","m²",1.1,310.0],["10.330.3501","2 mm kalınlıkta parke altı şiltesi","m²",1.01,6.0]]],
[670,"15.490.1002","Laminat parke döşeme kaplaması yapılması (AC3 Sınıf 23-31) (süpürgelik dahil)","m²",588.51,643.04,"İnce İnşaat","",[["10.170.1202","AC3 Sınıf 23-31 Laminat parke","m²",1.1,310.0],["10.330.3501","2 mm kalınlıkta parke altı şiltesi","m²",1.01,6.0],["10.100.1009","Marangoz ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.15,205.0],["10.170.1203","AC4 Sınıf 32 Laminat parke","m²",1.1,364.0],["10.330.3501","2 mm kalınlıkta parke altı şiltesi","m²",1.05,6.0],["10.100.1009","Marangoz ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.15,205.0]]],
[671,"15.490.1003","Laminat parke döşeme kaplaması yapılması (AC4 Sınıf 32) (süpürgelik dahil)","m²",663.06,663.06,"İnce İnşaat","",[["10.170.1203","AC4 Sınıf 32 Laminat parke","m²",1.1,364.0],["10.330.3501","2 mm kalınlıkta parke altı şiltesi","m²",1.05,6.0],["10.100.1009","Marangoz ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.15,205.0]]],
[672,"15.495.1001","Ahşaptan süpürgelik yapılması ve yerine konulması","m",141.94,141.94,"İnce İnşaat","",[["10.130.4501","Çam kerestesi (I.sınıf)","m³",0.004,10200.0],["19.100.1091","","Sa",0.005,6059.7],["10.100.1017","Dülger ustası","Sa",0.08,310.0],["10.100.1062","Düz işçi","Sa",0.08,205.0],["10.130.4509","Kayın kerestesi","m³",0.012,12000.0],["10.130.4509","Kayın kerestesi","m³",0.0006,12000.0]]],
[673,"15.500.1003","Düz merdiven küpeştesi yapılması ve yerine konulması","m",1133.16,1133.16,"İnce İnşaat","",[["10.130.4509","Kayın kerestesi","m³",0.012,12000.0],["10.130.4509","Kayın kerestesi","m³",0.0006,12000.0],["15.540.1014","Ahşap yüzeylerin verniklenmesi","m²",0.17,234.38],["19.100.1091","","Sa",0.05,6059.7],["10.100.1008","Doğramacı ustası","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.130.4509","Kayın kerestesi","m³",0.04,12000.0],["10.130.4509","Kayın kerestesi","m³",0.002,12000.0]]],
[674,"15.500.1004","Eğri merdiven küpeştesi yapılması ve yerine konulması","m",1931.9,1931.9,"İnce İnşaat","",[["10.130.4509","Kayın kerestesi","m³",0.04,12000.0],["10.130.4509","Kayın kerestesi","m³",0.002,12000.0],["15.540.1014","Ahşap yüzeylerin verniklenmesi","m²",0.17,234.38],["19.100.1091","","Sa",0.07,6059.7],["10.100.1008","Doğramacı ustası","Sa",1.4,310.0],["10.100.1062","Düz işçi","Sa",0.7,205.0]]],
[675,"15.505.1001","Ahşaptan lambri yapılması","m²",4524.55,4524.55,"İnce İnşaat","",[["10.130.4502","Çam kerestesi (II.sınıf)","m³",0.015,9000.0],["10.130.4509","Kayın kerestesi","m³",0.032,12000.0],["10.170.1801","Kontraplak","m³",0.0053,25790.0],["19.100.1091","","Sa",0.35,6059.7],["10.100.1008","Doğramacı ustası","Sa",2.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.4504","Beyaz çam kerestesi","m³",0.06,10200.0]]],
[676,"15.510.1001","Ahşaptan masif tablalı iç kapı kasa ve pervazı yapılması yerine konulması","m²",2322.39,2322.39,"İnce İnşaat","",[["10.130.4504","Beyaz çam kerestesi","m³",0.06,10200.0],["19.100.1091","","Sa",0.12,6059.7],["10.100.1008","Doğramacı ustası","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[677,"15.510.1002","Ahşaptan masif tablalı dış kapı kasa ve pervazı yapılması yerine konulması","m²",2524.89,2524.89,"İnce İnşaat","",[["10.130.4501","Çam kerestesi (I.sınıf)","m³",0.064,10200.0],["19.100.1091","","Sa",0.14,6059.7],["10.100.1008","Doğramacı ustası","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[678,"15.510.1101","Ahşaptan masif tablalı iç kapı kanadı yapılması ve yerine konulması","m²",2158.08,2158.08,"İnce İnşaat","1) Kapı kanadının dıştan dışa eni ve boyu çarpılarak alan hesaplanır. Bu ölçüye kapı kasaları dâhil edilmez.",[["10.130.4504","Beyaz çam kerestesi","m³",0.053,10200.0],["19.100.1091","","Sa",0.14,6059.7],["10.100.1017","Dülger ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[679,"15.510.1102","Ahşaptan masif tablalı dış kapı kanadı yapılması ve yerine konulması","m²",2284.83,2284.83,"İnce İnşaat","1) Kapı kanadının dıştan dışa eni ve boyu çarpılarak alan hesaplanır. Bu ölçüye kapı kasaları dâhil edilmez.",[["10.130.4501","Çam kerestesi (I.sınıf)","m³",0.057,10200.0],["19.100.1091","","Sa",0.15,6059.7],["10.100.1017","Dülger ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[680,"15.510.1105","Ahşaptan camlı çarpma iç kapı kanadı yapılması ve yerine konulması","m²",1894.08,1894.08,"İnce İnşaat","1) Kapı kanadının dıştan dışa eni ve boyu çarpılarak alan hesaplanır. Bu ölçüye kapı kasaları dâhil edilmez.",[["10.130.4501","Çam kerestesi (I.sınıf)","m³",0.025,10200.0],["10.420.1151","Prinç ağac vidası","adet",30.0,0.5],["19.100.1091","","Sa",0.15,6059.7],["10.100.1017","Dülger ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[681,"15.510.9991","Mevcut kapılara suni deri ile kapitone kaplama yapılması","m²",2794.88,2794.88,"İnce İnşaat","",[["10.420.1515","İyi cins suni deri","m²",1.1,40.0],["10.420.1516","Şerit kordon","m",3.5,6.5],["10.420.1153","Özel başlı çivi","adet",80.0,1.0],["10.330.5491","Kaput bezi","m²",1.1,7.5],["10.100.1025","Döşemeci ustası","Sa",4.0,310.0],["10.100.1062","Düz işçi","Sa",4.0,205.0],["10.130.4501","Çam kerestesi (I.sınıf)","m³",0.045,10200.0]]],
[682,"15.515.1001","Ahşaptan kasa ve pervazlı tek satıhlı pencere yapılması ve yerine konulması","m²",2123.95,2123.95,"İnce İnşaat","Kasaların 15 cm den fazla genişlikte olması halinde; 15 cm den fazla olan kasa genişliği alanı, pencere 1 m² fiyatının %",[["10.130.4501","Çam kerestesi (I.sınıf)","m³",0.045,10200.0],["19.100.1091","","Sa",0.15,6059.7],["10.100.1017","Dülger ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[683,"15.515.1101","Ahşaptan iç camekan yapılması ve yerine konulması","m²",2003.08,2003.08,"İnce İnşaat","Kasaların 15 cm den fazla genişlikte olması halinde; 15 cm den fazla olan kasa genişliği alanı, pencere 1 m² fiyatının %",[["10.130.4504","Beyaz çam kerestesi","m³",0.035,10200.0],["19.100.1091","","Sa",0.15,6059.7],["10.100.1008","Doğramacı ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[684,"15.525.1001","Ahşap çerçeveli, plastik telden sineklik yapılması ve yerine konulması (takılır-sökülür)","m²",1170.11,1170.11,"İnce İnşaat","",[["10.130.4501","Çam kerestesi (I.sınıf)","m³",0.015,10200.0],["10.420.1004","Sineklik teli","m²",1.1,55.0],["19.100.1091","","Sa",0.05,6059.7],["10.100.1008","Doğramacı ustası","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.420.1004","Sineklik teli","m²",1.1,55.0]]],
[685,"15.525.1002","Alüminyum çerçeveli, plastik telden sineklik yapılması ve yerine konulması (takılır-sökülür)","m²",1152.93,1152.93,"İnce İnşaat","",[["10.420.1004","Sineklik teli","m²",1.1,55.0],["19.100.1087","Alüminyum doğrama imalat atelyesi","Sa",0.05,4634.86],["10.100.1008","Doğramacı ustası","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[686,"15.525.1003","Pvc çerçeveli, plastik telden sineklik yapılması ve yerine konulması (takılır-sökülür)","m²",1123.46,1123.46,"İnce İnşaat","",[["10.420.1004","Sineklik teli","m²",1.1,55.0],["19.100.1088","Plastik doğrama imalat atelyesi","Sa",0.05,4603.42],["10.100.1008","Doğramacı ustası","Sa",1.0,310.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[687,"15.530.1151","oranı azaltılmış, kırılma dayanımı artırılmış alçı levhalar ile mevcut duvar üzeri, 60 cm aks","m²",1781.79,1781.79,"İnce İnşaat","",[["10.240.5723","","m²",1.05,129.0],["10.201.3029","","m",1.89,45.55],["10.201.3026","L 75 bağlantı elemanı, 2 mm kalınlıkta","adet",3.2,14.5],["10.420.1012","Vida ve plastik dübel","adet",4.0,2.0],["10.200.3137","Çelik dübel","adet",3.4,2.32],["10.100.1033","Alçı Levha Ustası","Sa",1.8,310.0],["10.100.1038","Alçı Levha Usta Yardımcısı","Sa",1.8,230.0],["10.100.1062","Düz işçi","Sa",0.8,205.0]]],
[688,"15.530.1152","oranı azaltılmış, kırılma dayanımı artırılmış alçı levhalar ile mevcut duvar üzeri, 60 cm aks","m²",1793.71,1793.71,"İnce İnşaat","",[["10.240.5723","","m²",1.05,129.0],["10.201.3006","Duvar C 75 profilli 0.90 mm kalınlıkta","m",1.89,50.6],["10.201.3026","L 75 bağlantı elemanı, 2 mm kalınlıkta","adet",3.2,14.5],["10.420.1012","Vida ve plastik dübel","adet",4.0,2.0],["10.200.3137","Çelik dübel","adet",3.4,2.32],["10.100.1033","Alçı Levha Ustası","Sa",1.8,310.0],["10.100.1038","Alçı Levha Usta Yardımcısı","Sa",1.8,230.0],["10.100.1062","Düz işçi","Sa",0.8,205.0]]],
[689,"15.530.1254","(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1153.59,1153.59,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.1,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.1,310.0]]],
[690,"15.530.1258","(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1289.83,1289.83,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.8875,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.25,310.0]]],
[691,"15.530.1262","(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1194.68,1194.68,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,84.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.1,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.1,310.0]]],
[692,"15.530.1266","(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1335.44,1335.44,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,84.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.8875,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.25,310.0]]],
[693,"15.530.1270","(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1197.46,1197.46,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,84.5],["10.200.3016","Duvar U 100 profilli 0.60 mm kalınlıkta","m",0.84,38.3],["10.200.3010","Duvar C 100 profilli 0.60 mm kalınlıkta","m",2.1,43.8],["10.200.3034","","m",1.3,3.69],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.1,310.0]]],
[694,"15.530.1274","(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1344.33,1344.33,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,84.5],["10.200.3016","Duvar U 100 profilli 0.60 mm kalınlıkta","m",0.84,38.3],["10.200.3010","Duvar C 100 profilli 0.60 mm kalınlıkta","m",2.8875,43.8],["10.200.3034","","m",1.3,3.69],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.25,310.0]]],
[695,"15.530.1304","(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1190.34,1190.34,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,98.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.1,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.1,310.0]]],
[696,"15.530.1308","(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1326.58,1326.58,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,98.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.8875,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.25,310.0]]],
[697,"15.530.1312","(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1231.43,1231.43,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,98.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.1,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.1,310.0]]],
[698,"15.530.1316","(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1372.19,1372.19,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,98.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.8875,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.25,310.0]]],
[699,"15.530.1320","(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1234.21,1234.21,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,98.5],["10.200.3016","Duvar U 100 profilli 0.60 mm kalınlıkta","m",0.84,38.3],["10.200.3010","Duvar C 100 profilli 0.60 mm kalınlıkta","m",2.1,43.8],["10.200.3034","","m",1.3,3.69],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.1,310.0]]],
[700,"15.530.1324","(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1381.08,1381.08,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,98.5],["10.200.3016","Duvar U 100 profilli 0.60 mm kalınlıkta","m",0.84,38.3],["10.200.3010","Duvar C 100 profilli 0.60 mm kalınlıkta","m",2.8875,43.8],["10.200.3034","","m",1.3,3.69],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.25,310.0]]],
[701,"15.530.1354","(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1514.15,1514.15,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.1,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.3,310.0]]],
[702,"15.530.1358","(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1652.89,1652.89,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.8875,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.45,310.0]]],
[703,"15.530.1362","(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1555.24,1555.24,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,84.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.1,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.3,310.0]]],
[704,"15.530.1366","(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1698.5,1698.5,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,84.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.8875,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.45,310.0]]],
[705,"15.530.1370","(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1558.03,1558.03,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,84.5],["10.200.3016","Duvar U 100 profilli 0.60 mm kalınlıkta","m",0.84,38.3],["10.200.3010","Duvar C 100 profilli 0.60 mm kalınlıkta","m",2.1,43.8],["10.200.3034","","m",1.3,3.69],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.3,310.0]]],
[706,"15.530.1374","(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1707.39,1707.39,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,84.5],["10.200.3016","Duvar U 100 profilli 0.60 mm kalınlıkta","m",0.84,38.3],["10.200.3010","Duvar C 100 profilli 0.60 mm kalınlıkta","m",2.8875,43.8],["10.200.3034","","m",1.3,3.69],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.45,310.0]]],
[707,"15.530.1404","(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1587.65,1587.65,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,98.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.1,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.3,310.0]]],
[708,"15.530.1408","(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1726.39,1726.39,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,98.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.8875,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.45,310.0]]],
[709,"15.530.1412","(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1628.74,1628.74,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,98.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.1,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.3,310.0]]],
[710,"15.530.1416","(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1772.0,1772.0,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,98.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.8875,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.45,310.0]]],
[711,"15.530.1420","(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1631.53,1631.53,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,98.5],["10.200.3016","Duvar U 100 profilli 0.60 mm kalınlıkta","m",0.84,38.3],["10.200.3010","Duvar C 100 profilli 0.60 mm kalınlıkta","m",2.1,43.8],["10.200.3034","","m",1.3,3.69],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.3,310.0]]],
[712,"15.530.1424","(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1780.89,1780.89,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,98.5],["10.200.3016","Duvar U 100 profilli 0.60 mm kalınlıkta","m",0.84,38.3],["10.200.3010","Duvar C 100 profilli 0.60 mm kalınlıkta","m",2.8875,43.8],["10.200.3034","","m",1.3,3.69],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.45,310.0]]],
[713,"15.530.1454","(her iki yüzünde üç kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1875.96,1875.96,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",6.3,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.1,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.5,310.0]]],
[714,"15.530.1458","(her iki yüzünde üç kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1917.05,1917.05,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",6.3,84.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.1,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.5,310.0]]],
[715,"15.530.1462","(her iki yüzünde üç kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1919.84,1919.84,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",6.3,84.5],["10.200.3016","Duvar U 100 profilli 0.60 mm kalınlıkta","m",0.84,38.3],["10.200.3010","Duvar C 100 profilli 0.60 mm kalınlıkta","m",2.1,43.8],["10.200.3034","","m",1.3,3.69],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.5,310.0]]],
[716,"15.530.1504","(her iki yüzünde iki kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",1940.74,1940.74,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",1.68,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",4.2,33.0],["10.200.3032","","m",2.6,1.38],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.75,310.0]]],
[717,"15.530.1508","(her iki yüzünde iki kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",2014.24,2014.24,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,98.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",1.68,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",4.2,33.0],["10.200.3032","","m",2.6,1.38],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.75,310.0]]],
[718,"15.530.1512","(her iki yüzünde iki kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı","m²",2022.89,2022.89,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",4.2,84.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",1.68,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",4.2,37.6],["10.200.3033","","m",2.6,2.75],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.75,310.0]]],
[719,"15.530.1551","(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm standart alçı levha","m²",1976.24,1976.24,"İnce İnşaat","",[["10.240.5583","","m²",4.3,45.0],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",1.68,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",4.2,33.0],["10.200.3032","","m",2.6,1.38],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[720,"15.530.1552","(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm su emme oranı","m²",2078.36,2078.36,"İnce İnşaat","",[["10.240.5603","12,5 mm kalınlığında (Tip H2) alçı levha","m²",4.3,64.0],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",1.68,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",4.2,33.0],["10.200.3032","","m",2.6,1.38],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[721,"15.530.1553","(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm yangına dayanımı","m²",2072.99,2072.99,"İnce İnşaat","",[["10.240.5613","12,5 mm kalınlığında (Tip F) alçı levha","m²",4.3,63.0],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",1.68,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",4.2,33.0],["10.200.3032","","m",2.6,1.38],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[722,"15.530.1554","(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm su emme oranı","m²",2188.55,2188.55,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",4.3,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",1.68,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",4.2,33.0],["10.200.3032","","m",2.6,1.38],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[723,"15.530.1555","(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 15 mm standart alçı levha ile)","m²",2051.49,2051.49,"İnce İnşaat","",[["10.240.5584","","m²",4.3,59.0],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",1.68,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",4.2,33.0],["10.200.3032","","m",2.6,1.38],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[724,"15.530.1556","(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 15 mm su emme oranı","m²",2175.11,2175.11,"İnce İnşaat","",[["10.240.5604","15 mm kalınlığında (Tip H2) alçı levha","m²",4.3,82.0],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",1.68,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",4.2,33.0],["10.200.3032","","m",2.6,1.38],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[725,"15.530.1557","(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 15 mm yangına dayanımı","m²",2140.18,2140.18,"İnce İnşaat","",[["10.240.5614","15 mm kalınlığında (Tip F) alçı levha","m²",4.3,75.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",1.68,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",4.2,33.0],["10.200.3032","","m",2.6,1.38],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[726,"15.530.1558","(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 15 mm su emme oranı","m²",2263.8,2263.8,"İnce İnşaat","",[["10.240.5634","15 mm kalınlığında (Tip FH2) alçı levha","m²",4.3,98.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",1.68,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",4.2,33.0],["10.200.3032","","m",2.6,1.38],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[727,"15.530.1559","(Duvar C 75 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm standart alçı levha","m²",2058.39,2058.39,"İnce İnşaat","",[["10.240.5583","","m²",4.3,45.0],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",1.68,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",4.2,37.6],["10.200.3033","","m",2.6,2.75],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[728,"15.530.1560","(Duvar C 75 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm su emme oranı","m²",2160.51,2160.51,"İnce İnşaat","",[["10.240.5603","12,5 mm kalınlığında (Tip H2) alçı levha","m²",4.3,64.0],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",1.68,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",4.2,37.6],["10.200.3033","","m",2.6,2.75],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[729,"15.530.1561","(Duvar C 75 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm yangına dayanımı","m²",2155.14,2155.14,"İnce İnşaat","",[["10.240.5613","12,5 mm kalınlığında (Tip F) alçı levha","m²",4.3,63.0],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",1.68,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",4.2,37.6],["10.200.3033","","m",2.6,2.75],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[730,"15.530.1562","(Duvar C 75 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm su emme oranı","m²",2270.7,2270.7,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",4.3,84.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",1.68,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",4.2,37.6],["10.200.3033","","m",2.6,2.75],["10.420.1012","Vida ve plastik dübel","adet",4.4,2.0],["10.200.3031","Derz bandı","m",3.0,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.001,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",2.1,310.0]]],
[731,"15.530.1754","aralığı) (12,5 mm çift kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha","m²",969.25,969.25,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,84.5],["10.200.3002","Tavan C 60 profilli 0.60 mm kalınlıkta","m",0.84,25.75],["10.200.3004","Tavan U 28 profilli 0.60 mm kalınlıkta","m",2.1,14.5],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3028","Agraf 12 cm","adet",1.0,4.55],["10.200.3031","Derz bandı","m",1.5,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.0005,4593.95]]],
[732,"15.530.1779","aralığı) (12,5 mm tek kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha","m²",900.21,900.21,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",1.05,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.1,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",1.5,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.0005,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",0.9,310.0]]],
[733,"15.530.1783","aralığı) (12,5 mm tek kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha","m²",1068.95,1068.95,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",1.05,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.8875,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",1.5,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.0005,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.1,310.0]]],
[734,"15.530.1787","aralığı) (12,5 mm tek kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha","m²",941.3,941.3,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",1.05,84.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.1,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",1.5,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.0005,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",0.9,310.0]]],
[735,"15.530.1791","aralığı) (12,5 mm tek kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha","m²",1114.56,1114.56,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",1.05,84.5],["10.200.3014","Duvar U 75 profilli 0.60 mm kalınlıkta","m",0.84,52.6],["10.200.3008","Duvar C 75 profilli 0.60 mm kalınlıkta","m",2.8875,37.6],["10.200.3033","","m",1.3,2.75],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",1.5,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.0005,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.1,310.0]]],
[736,"15.530.1829","aralığı) (12,5 mm çift kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha","m²",1148.11,1148.11,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.1,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",1.5,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.0005,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.1,310.0]]],
[737,"15.530.1833","aralığı) (12,5 mm çift kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha","m²",1183.1,1183.1,"İnce İnşaat","",[["10.240.5633","12,5 mm kalınlığında (Tip FH2) alçı levha","m²",2.1,84.5],["10.200.3012","Duvar U 50 profilli 0.60 mm kalınlıkta","m",0.84,27.1],["10.200.3006","Duvar C 50 profilli 0.60 mm kalınlıkta","m",2.8875,33.0],["10.200.3032","","m",1.3,1.38],["10.420.1012","Vida ve plastik dübel","adet",2.2,2.0],["10.200.3031","Derz bandı","m",1.5,0.86],["19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",0.0005,4593.95],["10.100.1033","Alçı Levha Ustası","Sa",1.1,310.0]]],
[738,"15.530.1901","900 mm, ana taşıyıcı profil mesafesi 1000 mm, tali taşıyıcı profil mesafesi 500 mm aks","m²",1226.94,1336.2,"İnce İnşaat","",[["10.240.5583","","m²",1.05,45.0],["10.200.3003","Tavan U 28 profilli 0.50 mm kalınlıkta","m",0.84,12.9],["10.200.3002","Tavan C 60 profilli 0.60 mm kalınlıkta","m",3.675,25.75],["10.200.3032","","m",0.8,1.38],["10.200.3026","Ekleme parçası","adet",0.7,3.04],["10.420.1012","Vida ve plastik dübel","adet",1.6,2.0],["10.200.3137","Çelik dübel","adet",1.7,2.32],["10.200.3028","Agraf 12 cm","adet",1.7,4.55]]],
[739,"15.530.1902","900 mm, ana taşıyıcı profil mesafesi 1000 mm, tali taşıyıcı profil mesafesi 500 mm aks","m²",1251.88,1251.88,"İnce İnşaat","",[["10.240.5603","12,5 mm kalınlığında (Tip H2) alçı levha","m²",1.05,64.0],["10.200.3003","Tavan U 28 profilli 0.50 mm kalınlıkta","m",0.84,12.9],["10.200.3002","Tavan C 60 profilli 0.60 mm kalınlıkta","m",3.675,25.75],["10.200.3032","","m",0.8,1.38],["10.200.3026","Ekleme parçası","adet",0.7,3.04],["10.420.1012","Vida ve plastik dübel","adet",1.6,2.0],["10.200.3137","Çelik dübel","adet",1.7,2.32],["10.200.3028","Agraf 12 cm","adet",1.7,4.55]]],
[740,"15.530.1903","900 mm, ana taşıyıcı profil mesafesi 1000 mm, tali taşıyıcı profil mesafesi 500 mm aks","m²",1250.56,1250.56,"İnce İnşaat","",[["10.240.5613","12,5 mm kalınlığında (Tip F) alçı levha","m²",1.05,63.0],["10.200.3003","Tavan U 28 profilli 0.50 mm kalınlıkta","m",0.84,12.9],["10.200.3002","Tavan C 60 profilli 0.60 mm kalınlıkta","m",3.675,25.75],["10.200.3032","","m",0.8,1.38],["10.200.3026","Ekleme parçası","adet",0.7,3.04],["10.420.1012","Vida ve plastik dübel","adet",1.6,2.0],["10.200.3137","Çelik dübel","adet",1.7,2.32],["10.200.3028","Agraf 12 cm","adet",1.7,4.55]]],
[741,"15.530.1905","yönde 900 mm, ana taşıyıcı profil mesafesi 1000 mm, tali taşıyıcı profil mesafesi 500 mm aks","m²",1234.16,1234.16,"İnce İnşaat","",[["10.240.5583","","m²",1.05,45.0],["10.200.3003","Tavan U 28 profilli 0.50 mm kalınlıkta","m",0.84,12.9],["10.200.3002","Tavan C 60 profilli 0.60 mm kalınlıkta","m",3.675,25.75],["10.200.3032","","m",0.8,1.38],["10.200.3026","Ekleme parçası","adet",0.7,3.04],["10.420.1012","Vida ve plastik dübel","adet",1.6,2.0],["10.200.3137","Çelik dübel","adet",1.7,2.32],["10.200.3132","Askı çubuğu 80cm","adet",1.7,2.58]]],
[742,"15.530.1906","yönde 900 mm, ana taşıyıcı profil mesafesi 1000 mm, tali taşıyıcı profil mesafesi 500 mm aks","m²",1259.1,1259.1,"İnce İnşaat","",[["10.240.5603","12,5 mm kalınlığında (Tip H2) alçı levha","m²",1.05,64.0],["10.200.3003","Tavan U 28 profilli 0.50 mm kalınlıkta","m",0.84,12.9],["10.200.3002","Tavan C 60 profilli 0.60 mm kalınlıkta","m",3.675,25.75],["10.200.3032","","m",0.8,1.38],["10.200.3026","Ekleme parçası","adet",0.7,3.04],["10.420.1012","Vida ve plastik dübel","adet",1.6,2.0],["10.200.3137","Çelik dübel","adet",1.7,2.32],["10.200.3132","Askı çubuğu 80cm","adet",1.7,2.58]]],
[743,"15.530.1907","yönde 900 mm, ana taşıyıcı profil mesafesi 1000 mm, tali taşıyıcı profil mesafesi 500 mm aks","m²",1257.79,1257.79,"İnce İnşaat","",[["10.240.5613","12,5 mm kalınlığında (Tip F) alçı levha","m²",1.05,63.0],["10.200.3003","Tavan U 28 profilli 0.50 mm kalınlıkta","m",0.84,12.9],["10.200.3002","Tavan C 60 profilli 0.60 mm kalınlıkta","m",3.675,25.75],["10.200.3032","","m",0.8,1.38],["10.200.3026","Ekleme parçası","adet",0.7,3.04],["10.420.1012","Vida ve plastik dübel","adet",1.6,2.0],["10.200.3137","Çelik dübel","adet",1.7,2.32],["10.200.3132","Askı çubuğu 80cm","adet",1.7,2.58]]],
[744,"15.530.1928","yönde 750 mm, ana taşıyıcı profil mesafesi 800 mm, tali taşıyıcı profil mesafesi 500 mm aks","m²",1474.03,1474.03,"İnce İnşaat","",[["10.240.5613","12,5 mm kalınlığında (Tip F) alçı levha","m²",2.1,63.0],["10.200.3003","Tavan U 28 profilli 0.50 mm kalınlıkta","m",0.84,12.9],["10.200.3002","Tavan C 60 profilli 0.60 mm kalınlıkta","m",3.675,25.75],["10.200.3032","","m",0.8,1.38],["10.200.3026","Ekleme parçası","adet",0.7,3.04],["10.420.1012","Vida ve plastik dübel","adet",1.6,2.0],["10.200.3137","Çelik dübel","adet",2.1,2.32],["10.200.3028","Agraf 12 cm","adet",2.1,4.55]]],
[745,"15.530.1932","yönde 750 mm, ana taşıyıcı profil mesafesi 800 mm, tali taşıyıcı profil mesafesi 500 mm aks","m²",1497.49,1497.49,"İnce İnşaat","",[["10.240.5613","12,5 mm kalınlığında (Tip F) alçı levha","m²",2.1,63.0],["10.200.3003","Tavan U 28 profilli 0.50 mm kalınlıkta","m",0.84,12.9],["10.200.3002","Tavan C 60 profilli 0.60 mm kalınlıkta","m",3.675,25.75],["10.200.3032","","m",0.8,1.38],["10.200.3026","Ekleme parçası","adet",0.7,3.04],["10.420.1012","Vida ve plastik dübel","adet",1.6,2.0],["10.200.3137","Çelik dübel","adet",2.1,2.32],["10.200.3132","Askı çubuğu 80cm","adet",2.1,2.58]]],
[746,"15.535.1001","(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 serisi) oturmalı sistem asma","m²",1098.74,1197.76,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6501","Alüminyum plaka","m²",1.05,480.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[747,"15.535.1002","boyalı(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) oturmalı sistem","m²",1124.99,1124.99,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6502","Alüminyum plaka","m²",1.05,500.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[748,"15.535.1003","(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000","m²",1216.86,1216.86,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6503","Alüminyum plaka","m²",1.05,570.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[749,"15.535.1004","(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 serisi) oturmalı sistem asma","m²",1517.86,1517.86,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6504","Alüminyum plaka","m²",1.05,530.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[750,"15.535.1005","(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 SERISI) oturmalı sistem asma","m²",1544.11,1544.11,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6505","Alüminyum plaka","m²",1.05,550.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[751,"15.535.1006","(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) oturmalı sistem asma","m²",1524.43,1524.43,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6506","Alüminyum plaka","m²",1.05,535.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[752,"15.535.1007","(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) oturmalı sistem asma","m²",1544.11,1544.11,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6507","Alüminyum plaka","m²",1.05,550.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[753,"15.535.1008","(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000","m²",1544.11,1544.11,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6508","Alüminyum plaka","m²",1.05,550.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[754,"15.535.1009","(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000","m²",1570.36,1570.36,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6509","Alüminyum plaka","m²",1.05,570.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[755,"15.535.1010","(polyester esaslı) deliksiz sıcak daldırma galvanize sac plakadan oturmalı sistem asma tavan","m²",908.43,908.43,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6510","Galvanize sac plaka","m²",1.05,335.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[756,"15.535.1011","(polyester esaslı) delikli sıcak daldırma galvanize sac plakadan oturmalı sistem asma tavan","m²",914.99,914.99,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6511","Galvanize sac plaka","m²",1.05,340.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[757,"15.535.1012","(polyester esaslı) arka yüzü akustik kumaş kaplı delikli sıcak daldrıma galvanize sac","m²",954.36,954.36,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6512","Galvanize sac plaka","m²",1.05,370.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3125","Kenar L Profili","m",1.1,11.35],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[758,"15.535.1013","(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 serisi) sarkmalı sistem asma","m²",1105.4,1105.4,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6551","Alüminyum plaka","m²",1.05,480.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[759,"15.535.1014","(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) sarkmalı sistem asma","m²",1131.65,1131.65,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6552","Alüminyum plaka","m²",1.05,500.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[760,"15.535.1015","(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000","m²",1197.28,1197.28,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6553","Alüminyum plaka","m²",1.05,550.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[761,"15.535.1016","(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 serisi)","m²",1524.53,1524.53,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6554","Alüminyum plaka","m²",1.05,530.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[762,"15.535.1017","(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 SERISI)","m²",1531.09,1531.09,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6555","Alüminyum plaka","m²",1.05,535.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[763,"15.535.1018","(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) sarkmalı sistem asma","m²",1524.53,1524.53,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6556","Alüminyum plaka","m²",1.05,530.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[764,"15.535.1019","(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) sarkmalı sistem asma","m²",1550.78,1550.78,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6557","Alüminyum plaka","m²",1.05,550.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[765,"15.535.1020","(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000","m²",1583.59,1583.59,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6558","Alüminyum plaka","m²",1.05,575.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[766,"15.535.1021","(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000","m²",1583.59,1583.59,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6559","Alüminyum plaka","m²",1.05,575.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",6.1,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[767,"15.535.1022","(polyester esaslı) deliksiz sıcak daldırma galvanize sac plakadan sarkmalı sistem asma tavan","m²",869.15,869.15,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6560","Galvanize sac plaka","m²",1.05,300.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[768,"15.535.1023","(polyester esaslı) delikli sıcak daldırma galvanize sac plakadan sarkmalı sistem asma tavan","m²",908.53,908.53,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6561","Galvanize sac plaka","m²",1.05,330.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[769,"15.535.1024","(polyester esaslı) arka yüzü akustik kumaş kaplı delikli sıcak daldırma galvanize sac","m²",954.46,954.46,"İnce İnşaat","1) Asma tavan yapılan yüzeyler ölçülür.",[["10.240.6562","Galvanize sac plaka","m²",1.05,365.0],["10.200.3071","T24 Ana taşıyıcı","m",1.0,15.8],["10.200.3091","T24 Ara taşıyıcı","m",2.6,22.8],["10.200.3127","Kenar Z Profili","m",1.1,16.2],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4],["10.200.3129","Askı çubuğu","adet",1.0,1.2],["10.420.1012","Vida ve plastik dübel","adet",2.0,2.0]]],
[770,"15.535.1025","(polyester esaslı) deliksiz alümünyum plakadan (EN AW 3000 serisi) gizli taşıyıcılı sistem","m²",1248.86,1248.86,"İnce İnşaat","",[["10.240.6601","Alüminyum plaka","m²",1.05,500.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",2.61,16.5],["10.200.3053","Gizli tasıyıcılı alüminyum kenar C profili","m",2.6,26.2],["10.200.3056","Birleşim klipsi","adet",1.47,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",2.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[771,"15.535.1026","(polyester esaslı) delikli alümünyum plakadan (EN AW 3000 serisi) gizli taşıyıcılı sistem","m²",1314.49,1314.49,"İnce İnşaat","",[["10.240.6602","Alüminyum plaka","m²",1.05,550.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",2.61,16.5],["10.200.3053","Gizli tasıyıcılı alüminyum kenar C profili","m",2.6,26.2],["10.200.3056","Birleşim klipsi","adet",1.47,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",2.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[772,"15.535.1027","(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alümünyum plakadan (EN AW 3000","m²",1419.49,1419.49,"İnce İnşaat","",[["10.240.6603","Alüminyum plaka","m²",1.05,630.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",2.61,16.5],["10.200.3053","Gizli tasıyıcılı alüminyum kenar C profili","m",2.6,26.2],["10.200.3056","Birleşim klipsi","adet",1.47,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",2.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[773,"15.535.1028","(polyester esaslı) deliksiz alümünyum plakadan (EN AW 3000 serisi) gizli taşıyıcılı sistem","m²",1590.35,1590.35,"İnce İnşaat","",[["10.240.6604","Alüminyum plaka","m²",1.05,530.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",4.1,16.5],["10.200.3053","Gizli tasıyıcılı alüminyum kenar C profili","m",1.1,26.2],["10.200.3056","Birleşim klipsi","adet",2.5,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",3.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[774,"15.535.1029","(polyester esaslı) deliksiz alümünyum plakadan (EN AW 3000 serisi) gizli taşıyıcılı sistem","m²",1682.23,1682.23,"İnce İnşaat","",[["10.240.6605","Alüminyum plaka","m²",1.05,600.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",4.1,16.5],["10.200.3053","Gizli tasıyıcılı alüminyum kenar C profili","m",1.1,26.2],["10.200.3056","Birleşim klipsi","adet",2.5,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",3.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[775,"15.535.1030","(polyester esaslı) delikli alümünyum plakadan (EN AW 3000 serisi) gizli taşıyıcılı sistem","m²",1603.48,1603.48,"İnce İnşaat","",[["10.240.6606","Alüminyum plaka","m²",1.05,540.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",4.1,16.5],["10.200.3053","Gizli tasıyıcılı alüminyum kenar C profili","m",1.1,26.2],["10.200.3056","Birleşim klipsi","adet",2.5,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",3.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[776,"15.535.1031","(polyester esaslı) delikli alümünyum plakadan (EN AW 3000 serisi) gizli taşıyıcılı sistem","m²",1715.04,1715.04,"İnce İnşaat","",[["10.240.6607","Alüminyum plaka","m²",1.05,625.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",4.1,16.5],["10.200.3053","Gizli tasıyıcılı alüminyum kenar C profili","m",1.1,26.2],["10.200.3056","Birleşim klipsi","adet",2.5,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",3.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[777,"15.535.1032","boyalı(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alümünyum plakadan (EN","m²",1616.6,1616.6,"İnce İnşaat","",[["10.240.6608","Alüminyum plaka","m²",1.05,550.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",4.1,16.5],["10.200.3053","Gizli tasıyıcılı alüminyum kenar C profili","m",1.1,26.2],["10.200.3056","Birleşim klipsi","adet",2.5,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",3.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[778,"15.535.1033","boyalı(polyester esaslı) deliksiz sıcak daldırma galvanize sac plakadan gizli taşıyıcılı sistem","m²",1000.9,1000.9,"İnce İnşaat","",[["10.240.6610","Galvanize sac plaka","m²",1.05,355.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",2.61,16.5],["10.200.3054","Gizli taşıyıcılı sistemin sac kenar C profili","m",1.1,20.0],["10.200.3056","Birleşim klipsi","adet",1.47,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",2.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[779,"15.535.1034","boyalı(polyester esaslı) delikli sıcak daldırma galvanize sac plakadan gizli taşıyıcılı sistem","m²",1007.46,1007.46,"İnce İnşaat","",[["10.240.6611","Galvanize sac plaka","m²",1.05,360.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",2.61,16.5],["10.200.3054","Gizli taşıyıcılı sistemin sac kenar C profili","m",1.1,20.0],["10.200.3056","Birleşim klipsi","adet",1.47,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",2.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[780,"15.535.1035","(polyester esaslı) arka yüzü akustik kumaş kaplı delikli sıcak daldırma galvaniz sac","m²",1053.4,1053.4,"İnce İnşaat","",[["10.240.6612","Galvanize sac plaka","m²",1.05,395.0],["10.200.3052","Gizli tasıyıcılı profil (clip-in sistem)","m",2.61,16.5],["10.200.3054","Gizli taşıyıcılı sistemin sac kenar C profili","m",1.1,20.0],["10.200.3056","Birleşim klipsi","adet",1.47,3.0],["10.200.3055","Taşıyıcı ek parçası","adet",0.47,2.5],["10.200.3057","Bastırma klipsi","adet",2.0,2.82],["10.200.3137","Çelik dübel","adet",1.0,2.32],["10.200.3136","Çiftli yay","adet",1.0,2.4]]],
[781,"15.540.1014","Ahşap yüzeylerin verniklenmesi","m²",292.98,315.46,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.45,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[782,"15.540.1015","Ahşap yüzeylerin vernikli renkli ahşap koruyucu ile verniklenmesi","m²",300.88,300.88,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.45,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.45,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[783,"15.540.1016","Ahşap yüzeylerin renkli ahşap koruyucu ile korunması","m²",273.25,273.25,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.45,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.65,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0]]],
[784,"15.540.1017","Her cins ahşap parkenin cilalanması","m²",425.94,425.94,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.65,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.300.1602","Zımpara kağıdı","adet",1.0,8.5]]],
[785,"15.540.1021","Ahşap yüzeylere bir kat sentetik boya yapılması","m²",237.35,237.35,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.15,205.0]]],
[786,"15.540.1022","Ahşap yüzeylere bir kat sentetik esaslı mat boya yapılması","m²",231.6,231.6,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.3,310.0],["10.100.1062","Düz işçi","Sa",0.15,205.0]]],
[787,"15.540.1023","Ahşap yüzeylere iki kat sentetik esaslı parlak boya yapılması","m²",309.54,309.54,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[788,"15.540.1024","Ahşap yüzeylere iki kat sentetik esaslı mat boya yapılması","m²",298.04,298.04,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[789,"15.540.1111","Demir yüzeylere korozyona karşı iki kat antipas boya yapılması","m²",270.38,293.2,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.45,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[790,"15.540.1121","Demir yüzeylere iki kat antipas, iki kat sentetik esaslı parlak boya yapılması","m²",375.38,404.54,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0]]],
[791,"15.540.1122","Demir yüzeylere iki kat antipas, iki kat sentetik esaslı mat boya yapılması","m²",330.55,330.55,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[792,"15.540.1123","Demir yüzeylere iki kat solvent bazlı epoksi boya yapılması","m²",408.36,408.36,"İnce İnşaat","",[["10.100.1023","Boyacı ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0]]],
[793,"15.540.1421","Beton, sıva ve benzeri yüzeylere 1,5 mm kalınlıkta akrilik esaslı renkli kaplama yapılması","m²",567.19,611.23,"İnce İnşaat","",[["10.100.1023","Boyacı ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0],["10.100.1023","Boyacı ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[794,"15.540.1422","Beton, sıva ve benzeri yüzeylere 2 mm kalınlıkta akrilik esaslı renkli kaplama yapılması","m²",629.69,629.69,"İnce İnşaat","",[["10.100.1023","Boyacı ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0],["10.100.1023","Boyacı ustası","Sa",0.9,310.0],["10.100.1062","Düz işçi","Sa",0.45,205.0]]],
[795,"15.540.1423","Beton, sıva ve benzeri yüzeylere 3 mm kalınlıkta akrilik esaslı renkli kaplama yapılması","m²",731.25,731.25,"İnce İnşaat","",[["10.100.1023","Boyacı ustası","Sa",0.9,310.0],["10.100.1062","Düz işçi","Sa",0.45,205.0]]],
[796,"15.540.1427","Beton, sıva ve benzeri yüzeylere 1,5 mm kalınlıkta çimento esaslı kaplama yapılması","m²",465.38,465.38,"İnce İnşaat","",[["10.130.9991","Su","m³",0.01,55.0],["10.100.1023","Boyacı ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0],["10.130.9991","Su","m³",0.014,55.0]]],
[797,"15.540.1428","Beton, sıva ve benzeri yüzeylere 2 mm kalınlıkta çimento esaslı kaplama yapılması","m²",469.15,469.15,"İnce İnşaat","",[["10.130.9991","Su","m³",0.014,55.0],["10.100.1023","Boyacı ustası","Sa",0.8,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0],["10.130.9991","Su","m³",0.018,55.0],["10.100.1023","Boyacı ustası","Sa",0.9,310.0],["10.100.1062","Düz işçi","Sa",0.45,205.0]]],
[798,"15.540.1429","Beton, sıva ve benzeri yüzeylere 3 mm kalınlıkta çimento esaslı kaplama yapılması","m²",543.74,543.74,"İnce İnşaat","",[["10.130.9991","Su","m³",0.018,55.0],["10.100.1023","Boyacı ustası","Sa",0.9,310.0],["10.100.1062","Düz işçi","Sa",0.45,205.0],["10.300.1602","Zımpara kağıdı","adet",1.0,8.5]]],
[799,"15.540.1501","Brüt beton yüzeylerin alçı veya sıvaya hazırlanması (iç cephe)","m²",134.06,134.06,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0],["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0]]],
[800,"15.540.1502","Lekeli ve isli duvar yüzeylerin boya işlemine hazır hale getirilmesi (iç cephe)","m²",276.88,276.88,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.2,310.0],["10.100.1062","Düz işçi","Sa",0.1,205.0],["10.130.6002","Sönmüş kireç CL 80S","ton",0.00025,3300.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1023","Boyacı ustası","Sa",0.5,310.0]]],
[801,"15.540.1503","Eski boyalı yüzeylere üç kat beyaz kireç badana yapılması (iç cephe)","m²",259.98,259.98,"İnce İnşaat","",[["10.130.6002","Sönmüş kireç CL 80S","ton",0.00025,3300.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1023","Boyacı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.130.6002","Sönmüş kireç CL 80S","ton",0.00025,3300.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1023","Boyacı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[802,"15.540.1504","Eski boyalı yüzeylere üç kat renkli kireç badana yapılması (iç cephe)","m²",260.73,260.73,"İnce İnşaat","",[["10.130.6002","Sönmüş kireç CL 80S","ton",0.00025,3300.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1023","Boyacı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[803,"15.540.1515","Yeni sıva yüzeylere üç kat beyaz kireç badana yapılması (iç cephe)","m²",258.99,258.99,"İnce İnşaat","",[["10.130.6002","Sönmüş kireç CL 80S","ton",0.00025,3300.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1023","Boyacı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.130.6002","Sönmüş kireç CL 80S","ton",0.00025,3300.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1023","Boyacı ustası","Sa",0.5,310.0]]],
[804,"15.540.1516","Yeni sıva yüzeylere üç kat renkli kireç badana yapılması (iç cephe)","m²",259.74,259.74,"İnce İnşaat","",[["10.130.6002","Sönmüş kireç CL 80S","ton",0.00025,3300.0],["10.130.9991","Su","m³",0.002,55.0],["10.100.1023","Boyacı ustası","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.300.1602","Zımpara kağıdı","adet",1.0,8.5]]],
[805,"15.540.1532","Antibakteriyel İç Cephe Boyası (Örtücülük Sınıf:2, YOD:Sınıf2, Parlaklık:G3) yapılması (iç","m²",376.6,376.6,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.55,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0]]],
[806,"15.540.1533","Antibakteriyel İç Cephe Boyası (Örtücülük Sınıf:2, YOD:Sınıf1, Parlaklık:G2) yapılması (iç","m²",384.88,384.88,"İnce İnşaat","",[["10.300.1602","Zımpara kağıdı","adet",1.0,8.5],["10.100.1023","Boyacı ustası","Sa",0.55,310.0],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.300.1602","Zımpara kağıdı","adet",1.0,8.5]]],
[807,"15.540.1538","Antibakteriyel İç Cephe Boyası (Örtücülük Sınıf:2, YOD:Sınıf2, Parlaklık:G3) yapılması (iç","m²",240.09,240.09,"İnce İnşaat","",[["10.100.1023","Boyacı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[808,"15.540.1539","Antibakteriyel İç Cephe Boyası (Örtücülük Sınıf:2, YOD:Sınıf1, Parlaklık:G2) yapılması (iç","m²",247.44,247.44,"İnce İnşaat","",[["10.100.1023","Boyacı ustası","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.2,205.0]]],
[809,"15.540.1543","Yeni sıva yüzeylere astar uygulanarak iki kat sentetik esaslı parlak boya yapılması (iç cephe)","m²",371.19,371.19,"İnce İnşaat","",[["10.100.1023","Boyacı ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["10.100.1023","Boyacı ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0]]],
[810,"15.540.1544","Yeni sıva yüzeylere astar uygulanarak iki kat sentetik esaslı mat boya yapılması (iç cephe)","m²",357.99,357.99,"İnce İnşaat","",[["10.100.1023","Boyacı ustası","Sa",0.6,310.0],["10.100.1062","Düz işçi","Sa",0.3,205.0]]],
[811,"15.555.1001","daldırma galvaniz üzeri elektrostatik polyester toz boyalı panel teller ile çit yapılması (Direk","m",693.2,693.2,"Kaba İnşaat","",[["10.480.1501","","m",1.0,220.0],["10.480.1511","1,00 m yükseklikte panel çit direği","adet",0.4,195.0],["10.480.1521","Klips (panel çit için)","adet",0.8,8.0],["10.200.4024","","adet",1.65,4.35],["19.100.1110","Matkap","Sa",0.1,369.82],["10.100.1068","Birinci sınıf usta","Sa",0.4,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[812,"15.555.1002","daldırma galvaniz üzeri elektrostatik polyester toz boyalı panel teller ile çit yapılması (Direk","m",810.95,810.95,"Kaba İnşaat","",[["10.480.1502","","m",1.0,270.0],["10.480.1512","1,20 m yükseklikte panel çit direği","adet",0.4,220.0],["10.480.1521","Klips (panel çit için)","adet",1.2,8.0],["10.200.4024","","adet",1.65,4.35],["19.100.1110","Matkap","Sa",0.1,369.82],["10.100.1068","Birinci sınıf usta","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[813,"15.555.1003","daldırma galvaniz üzeri elektrostatik polyester toz boyalı panel teller ile çit yapılması (Direk","m",917.2,917.2,"Kaba İnşaat","",[["10.480.1503","","m",1.0,335.0],["10.480.1513","1,50 m yükseklikte panel çit direği","adet",0.4,270.0],["10.480.1521","Klips (panel çit için)","adet",1.2,8.0],["10.200.4024","","adet",1.65,4.35],["19.100.1110","Matkap","Sa",0.1,369.82],["10.100.1068","Birinci sınıf usta","Sa",0.5,310.0],["10.100.1062","Düz işçi","Sa",0.4,205.0]]],
[814,"19.100.1001","Ekskavatör'ün 1 saatlik ücreti","Sa",2183.46,2825.62,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[815,"19.100.1002","Ekskavatör Beko'nun 1 saatlik ücreti","Sa",2571.91,3355.18,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[816,"19.100.1003","Ekskavatör'ün 1 saatlik ücreti","Sa",2645.46,2645.46,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[817,"19.100.1004","Ekskavatör'ün 1 saatlik ücreti","Sa",2867.81,2867.81,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[818,"19.100.1005","Ekskavatör'ün 1 saatlik ücreti","Sa",3374.35,3374.35,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.12,230.0]]],
[819,"19.100.1006","Ekskavatör (paletli) (210-259 HP) (maksimum 2,5 m³) 1 saatlik ücreti","Sa",2971.46,2971.46,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.12,230.0],["10.100.1060","Formen","Sa",0.48,445.0]]],
[820,"19.100.1007","Ekskavatör'ün 1 saatlik ücreti","Sa",3836.26,3836.26,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.12,230.0]]],
[821,"19.100.1008","Ekskavatör (paletli) (260-299 HP) (maksimum 2,5 m³) 1 saatlik ücreti","Sa",3421.44,3421.44,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.12,230.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.12,230.0]]],
[822,"19.100.1009","Ekskavatör (paletli) (300-329 HP) (maksimum 3,5 m³) 1 saatlik ücreti","Sa",3984.58,3984.58,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.12,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0]]],
[823,"19.100.1010","Traktör Skreyper'in 1 saatlik ücreti","Sa",2198.58,2198.58,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[824,"19.100.1011","Traktör Ripper'in 1 saatlik ücreti","Sa",3389.66,3389.66,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[825,"19.100.1012","Motor Greyder'in 1 saatlik ücreti","Sa",1830.41,2345.87,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0]]],
[826,"19.100.1013","Greyder (190-209 HP) 1 saatlik ücreti","Sa",2622.44,2622.44,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0]]],
[827,"19.100.1014","Greyder (210-230 HP) 1 saatlik ücreti","Sa",2933.93,2933.93,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[828,"19.100.1015","Lastik tekerlekli traktör Skreyper'in 1 saatlik ücreti","Sa",4947.48,4947.48,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[829,"19.100.1016","Traktör Buldozer'in (70 HP) 1 saatlik ücreti","Sa",1789.23,2264.61,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[830,"19.100.1017","Traktör Buldozer'in (100 HP) 1 saatlik ücreti","Sa",2049.13,2686.47,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[831,"19.100.1018","Traktör Buldozer'in (160 HP) 1 saatlik ücreti","Sa",2586.03,2586.03,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[832,"19.100.1019","Traktör Buldozer'in 1 saatlik ücreti","Sa",3141.71,3141.71,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[833,"19.100.1020","Traktör Buldozer'in (285 HP) 1 saatlik ücreti","Sa",4868.73,4868.73,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[834,"19.100.1021","Traktör Buldozer'in (345 HP) 1 saatlik ücreti","Sa",5439.9,5439.9,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[835,"19.100.1022","Şahmerdan'ın (50 HP) 1 saatlik ücreti","Sa",2540.26,2540.26,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1054","Makinist","Sa",2.4,310.0],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",1.0,310.0],["10.100.1063","Erbab işçi","Sa",4.0,250.0]]],
[836,"19.100.1023","Kompresör'ün 1 saatlik ücreti","Sa",2508.08,2508.08,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",2.4,310.0],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",1.0,310.0],["10.100.1063","Erbab işçi","Sa",4.0,250.0],["10.100.1054","Makinist","Sa",2.4,310.0],["10.100.1063","Erbab işçi","Sa",1.0,250.0]]],
[837,"19.100.1024","Vantilatasyon için Kompresör'ün 1 saatlik ücreti","Sa",1461.78,1461.78,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",2.4,310.0],["10.100.1063","Erbab işçi","Sa",1.0,250.0],["10.100.1054","Makinist","Sa",1.2,310.0]]],
[838,"19.100.1025","Kompresör'ün 1 saatlik ücreti (250 HP)","Sa",1942.71,1942.71,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.2,310.0],["10.100.1054","Makinist","Sa",2.4,310.0],["10.100.1056","Makinist yardımcısı","Sa",1.0,260.0]]],
[839,"19.100.1026","Enjeksiyon Makinası'nın 1 saatlik ücreti","Sa",1471.78,1471.78,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",2.4,310.0],["10.100.1056","Makinist yardımcısı","Sa",1.0,260.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.12,230.0]]],
[840,"19.100.1027","Kazıcı yükleyici (100 HP) (maksimum 2,5 m³) 1 saatlik ücreti","Sa",1656.24,1656.24,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.12,230.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0]]],
[841,"19.100.1028","Lastik tekerlekli yükleyicinin 1 saatlik ücreti","Sa",1686.26,2196.53,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.12,230.0]]],
[842,"19.100.1029","Yükleyici (lastik tekerlekli) (100 HP) (maksimum 2 m³) 1 saatlik ücreti","Sa",1510.89,2015.26,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.12,230.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[843,"19.100.1030","Paletli Yükleyicinin 1 saatlik ücreti","Sa",2037.39,2037.39,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1054","Makinist","Sa",2.88,310.0]]],
[844,"19.100.1031","Betoniyer'in 1 saatlik ücreti","Sa",1012.75,1157.05,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",2.88,310.0],["10.420.1512","Mozaik silme taşı","adet",0.04,90.0],["10.130.9991","Su","m³",0.2,55.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1039","Mozayikçi usta yardımcısı","Sa",1.0,230.0]]],
[845,"19.100.1032","Mozayik silme makinası'nın 1 saatlik ücreti","Sa",528.95,528.95,"Makine/İşçilik","",[["10.420.1512","Mozaik silme taşı","adet",0.04,90.0],["10.130.9991","Su","m³",0.2,55.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1039","Mozayikçi usta yardımcısı","Sa",1.0,230.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0]]],
[846,"19.100.1033","Beton vibratörü'nün 1 saatlik ücreti","Sa",458.72,458.72,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0]]],
[847,"19.100.1034","Konkasör'ün 1 saatlik ücreti","Sa",4659.08,4659.08,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1054","Makinist","Sa",2.88,310.0],["10.100.1059","Yağcı","Sa",4.0,230.0],["10.100.1062","Düz işçi","Sa",3.0,205.0],["10.100.1060","Formen","Sa",0.24,445.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",0.5,230.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[848,"19.100.1035","Elek makinası'nın 1 saatlik ücreti","Sa",1334.95,1334.95,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.24,445.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",0.5,230.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[849,"19.100.1036","Eleme makinasının 1 saatlik ücreti","Sa",1254.75,1254.75,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.6,445.0],["10.100.1054","Makinist","Sa",1.2,310.0],["10.100.1059","Yağcı","Sa",0.12,230.0],["10.100.1062","Düz işçi","Sa",0.6,205.0],["10.100.1056","Makinist yardımcısı","Sa",1.0,260.0]]],
[850,"19.100.1037","Motopomp'un (5 Ps.) 1 saatlik ücreti","Sa",305.51,305.51,"Makine/İşçilik","",[["10.100.1056","Makinist yardımcısı","Sa",1.0,260.0],["10.100.1056","Makinist yardımcısı","Sa",1.2,260.0]]],
[851,"19.100.1038","Motopompun 1 saatlik ücreti","Sa",364.62,364.62,"Makine/İşçilik","",[["10.100.1056","Makinist yardımcısı","Sa",1.2,260.0],["10.100.1056","Makinist yardımcısı","Sa",1.0,260.0]]],
[852,"19.100.1039","Motopomp'un (20 Ps.) 1 saatlik ücreti","Sa",444.43,444.43,"Makine/İşçilik","",[["10.100.1056","Makinist yardımcısı","Sa",1.0,260.0],["10.100.1054","Makinist","Sa",1.44,310.0]]],
[853,"19.100.1040","Motopomp'un (30 Ps.) 1 saatlik ücreti","Sa",735.16,735.16,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1054","Makinist","Sa",1.44,310.0]]],
[854,"19.100.1041","Motopomp'un (45 Ps.) 1 saatlik ücreti","Sa",874.78,874.78,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1054","Makinist","Sa",1.44,310.0]]],
[855,"19.100.1042","Motopomp'un (60 Ps.) 1 saatlik ücreti","Sa",1004.89,1004.89,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1055","Operatör makinist","Sa",1.0,355.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[856,"19.100.1043","Çekilir tip beton pompası'nın 1 saatlik ücreti","Sa",1697.1,1697.1,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.0,355.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1051","Şöför","Sa",1.44,310.0],["10.100.1051","Şöför","Sa",1.44,310.0]]],
[857,"19.100.1044","Arazöz'ün 1 saatlik ücreti","Sa",882.41,882.41,"Makine/İşçilik","",[["10.100.1051","Şöför","Sa",1.44,310.0],["10.100.1051","Şöför","Sa",1.44,310.0]]],
[858,"19.100.1045","Arazöz Pick-Up'ın 1 saatlik ücreti","Sa",845.99,845.99,"Makine/İşçilik","",[["10.100.1051","Şöför","Sa",1.44,310.0],["10.100.1054","Makinist","Sa",2.88,310.0]]],
[859,"19.100.1046","Her cins (Titreşimli Darbeli) motorlu kompaktör'ün 1 saatlik ücreti","Sa",978.34,1110.65,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",2.88,310.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[860,"19.100.1047","Titreşimli Silindir'in 1 saatlik ücreti","Sa",1531.16,1531.16,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[861,"19.100.1048","Titreşimli Silindir'in 1 saatlik ücreti","Sa",1597.17,1597.17,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[862,"19.100.1049","Komple keçiayağı silindir'in 1 saatlik ücreti","Sa",1162.88,1162.88,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["19.100.1049","Komple keçiayağı silindir (40 HP)","Sa",1.6,1162.88],["19.100.1049","Komple keçiayağı silindir (40 HP)","Sa",2.1,1162.88]]],
[863,"19.100.1050","2 çift Tanburlu keçiayağı silindir'in 1 saatlik ücreti","Sa",1860.61,1860.61,"Makine/İşçilik","",[["19.100.1049","Komple keçiayağı silindir (40 HP)","Sa",1.6,1162.88],["19.100.1049","Komple keçiayağı silindir (40 HP)","Sa",2.1,1162.88],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[864,"19.100.1051","3 çift Tanburlu keçiayağı silindir'in 1 saatlik ücreti","Sa",2442.05,2442.05,"Makine/İşçilik","",[["19.100.1049","Komple keçiayağı silindir (40 HP)","Sa",2.1,1162.88],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[865,"19.100.1052","Demir merdaneli silindir'in 1 saatlik ücreti","Sa",1162.88,1162.88,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[866,"19.100.1053","Demir merdaneli silindir'in 1 saatlik ücreti","Sa",1324.45,1324.45,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[867,"19.100.1054","Lastik tekerlekli silindir'in 1 saatlik ücreti","Sa",1162.88,1162.88,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[868,"19.100.1055","Küçük eleme tesisi'nin 1 saatlik ücreti","Sa",1320.85,1320.85,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1054","Makinist","Sa",1.2,310.0],["10.100.1059","Yağcı","Sa",0.2,230.0]]],
[869,"19.100.1056","Fore kazık delgi makinası'nın (300 HP) 1 saatlik ücreti","Sa",7033.33,7033.33,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.2,310.0],["10.100.1059","Yağcı","Sa",0.2,230.0],["10.100.1054","Makinist","Sa",1.2,310.0],["10.100.1059","Yağcı","Sa",0.2,230.0]]],
[870,"19.100.1057","Fore kazık delgi makinası'nın (200 HP) 1 saatlik ücreti","Sa",3126.89,3126.89,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.2,310.0],["10.100.1059","Yağcı","Sa",0.2,230.0],["10.100.1054","Makinist","Sa",1.2,310.0],["10.100.1059","Yağcı","Sa",0.2,230.0]]],
[871,"19.100.1058","Fore kazık delgi makinası'nın (440 HP) 1 saatlik ücreti","Sa",8929.55,8929.55,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.2,310.0],["10.100.1059","Yağcı","Sa",0.2,230.0],["10.100.1055","Operatör makinist","Sa",1.0,355.0],["10.100.1057","Operatör yardımcısı","Sa",1.0,300.0]]],
[872,"19.100.1059","Otomatik Beton Santrali'nin (1000 Litre kapasiteli, 50 m³/saat) 1 saatlik ücreti","Sa",577.13,577.13,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.0,355.0],["10.100.1057","Operatör yardımcısı","Sa",1.0,300.0]]],
[873,"19.100.1060","Rotor sistemli BEP 80 M ve benzeri beton pompası'nın 1 saatlik ücreti","Sa",4205.04,4205.04,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.0,355.0],["10.100.1057","Operatör yardımcısı","Sa",1.0,300.0],["10.100.1054","Makinist","Sa",0.96,310.0]]],
[874,"19.100.1061","Kaynak Makinası'nın 1 saatlik ücreti","Sa",492.91,492.91,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",0.96,310.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1064","Çırak","Sa",1.0,205.0],["10.100.1519","Gemici","Sa",1.44,260.0]]],
[875,"19.100.1062","5 kw Jeneratör grubunun 1 saatlik ücreti","Sa",757.34,909.15,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1064","Çırak","Sa",1.0,205.0],["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0]]],
[876,"19.100.1063","Klepeli Taş Dubası'nın (125 Ton) 1 saatlik ücreti","Sa",610.65,610.65,"Makine/İşçilik","",[["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0]]],
[877,"19.100.1064","Klepeli Taş Dubası'nın (400 Ton) 1 saatlik ücreti","Sa",899.4,899.4,"Makine/İşçilik","",[["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0]]],
[878,"19.100.1065","DevrilirTaş Dubası'ının (300 Ton) 1 saatlik ücreti","Sa",899.4,899.4,"Makine/İşçilik","",[["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0]]],
[879,"19.100.1066","Klapeli Kum Dubası'ının (300 Ton) 1 saatlik ücreti","Sa",899.4,899.4,"Makine/İşçilik","",[["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0]]],
[880,"19.100.1067","Klapeli Kum Dubası'ının (2x255 HP, 500 m³) 1 saatlik ücreti","Sa",2399.4,2399.4,"Makine/İşçilik","",[["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0]]],
[881,"19.100.1068","Motorsuz Layter'in (180 m³) 1 saatlik ücreti","Sa",1324.4,1324.4,"Makine/İşçilik","",[["10.100.1519","Gemici","Sa",1.44,260.0],["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[882,"19.100.1069","Romorkör'ün (116 HP) 1 saatlik ücreti","Sa",3336.64,3336.64,"Makine/İşçilik","",[["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[883,"19.100.1070","Romorkör'ün (240 HP) 1 saatlik ücreti","Sa",4544.99,4544.99,"Makine/İşçilik","",[["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[884,"19.100.1071","Romorkör'ün (310 HP) 1 saatlik ücreti","Sa",4989.42,4989.42,"Makine/İşçilik","",[["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0]]],
[885,"19.100.1072","Romorkör'ün (525 HP) 1 saatlik ücreti","Sa",7689.3,7689.3,"Makine/İşçilik","",[["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[886,"19.100.1073","Romorkör'ün (2x300 HP) 1 saatlik ücreti","Sa",8262.79,8262.79,"Makine/İşçilik","",[["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1513","Usta gemici","Sa",2.88,290.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1508","Maçula reisi","Sa",1.44,395.0],["10.100.1513","Usta gemici","Sa",4.32,290.0],["10.100.1514","Gemi yağcısı","Sa",2.88,290.0]]],
[887,"19.100.1074","Maçula'nın 1 saatlik ücreti","Sa",9418.22,9418.22,"Makine/İşçilik","",[["10.100.1504","Romorkör kaptanı","Sa",1.44,420.0],["10.100.1508","Maçula reisi","Sa",1.44,395.0],["10.100.1513","Usta gemici","Sa",4.32,290.0],["10.100.1514","Gemi yağcısı","Sa",2.88,290.0],["10.100.1511","Güverte lostromosu","Sa",2.88,300.0],["10.100.1505","Romorkör makinisti (Çarkçı)","Sa",1.44,420.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[888,"19.100.1075","Çim biçme makinası'nın 1 saatlik ücreti","Sa",206.49,206.49,"Makine/İşçilik","",[["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[889,"19.100.1076","Ziraat işlerinde kullanılan Traktör'ün 1 saatlik ücreti","Sa",1068.47,1068.47,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[890,"19.100.1077","Ziraat işlerinde kullanılan Traktör'ün 1 Saatlik Ücreti","Sa",1259.25,1259.25,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0]]],
[891,"19.100.1078","10 litrelik kollu sırt pülverizatör'ün 1 saatlik ücreti","Sa",275.65,275.65,"Makine/İşçilik","",[["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[892,"19.100.1079","10 litrelik motorlu sırt pülverizatör'ün 1 saatlik ücreti","Sa",314.69,314.69,"Makine/İşçilik","",[["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",2.0,205.0],["19.100.1045","Arazöz Pick-Up","Sa",1.0,845.99]]],
[893,"19.100.1080","El ile çekilen 100 litrelik motorlu pülverizatör'ün 1 saatlik ücreti","Sa",561.5,561.5,"Makine/İşçilik","",[["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",2.0,205.0],["19.100.1045","Arazöz Pick-Up","Sa",1.0,845.99],["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0]]],
[894,"19.100.1081","Araç ile çekilen 250 litrelik motorlu pülverizatör'ün 1 saatlik ücreti","Sa",1651.68,1651.68,"Makine/İşçilik","",[["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",2.0,205.0],["19.100.1045","Arazöz Pick-Up","Sa",1.0,845.99],["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",2.0,205.0],["19.100.1045","Arazöz Pick-Up","Sa",1.0,845.99]]],
[895,"19.100.1082","Araç ile taşınır 560 litrelik motorlu pülverizatör'ün 1 saatlik ücreti","Sa",1768.39,1768.39,"Makine/İşçilik","",[["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",2.0,205.0],["19.100.1045","Arazöz Pick-Up","Sa",1.0,845.99]]],
[896,"19.100.1083","Kum Kutusu ve Nozul'un 1 saatlik ücreti","Sa",26.77,26.77,"Makine/İşçilik","",[["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1051","Şöför","Sa",1.0,310.0]]],
[897,"19.100.1084","Kendi yürür 1200 litrelik motorlu pülverizatör'ün 1 saatlik ücreti","Sa",1150.38,1150.38,"Makine/İşçilik","",[["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1051","Şöför","Sa",1.0,310.0],["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",4.0,205.0],["19.100.1044","Arazöz","Sa",1.0,882.41]]],
[898,"19.100.1085","Karıştırıcı (Blender)'in 1 saatlik ücreti","Sa",7.59,7.59,"Makine/İşçilik","",[["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",4.0,205.0],["19.100.1044","Arazöz","Sa",1.0,882.41],["10.130.9991","Su","m³",0.5,55.0]]],
[899,"19.100.1086","Araç ile taşınır hidrolik düzenli 2200 litrelik motorlu pülverizatör'ün 1 saatlik ücreti","Sa",2290.4,2290.4,"Makine/İşçilik","",[["10.100.1072","Pulverizatör operatörü","Sa",1.0,275.0],["10.100.1062","Düz işçi","Sa",4.0,205.0],["19.100.1044","Arazöz","Sa",1.0,882.41],["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",1.0,445.0],["10.100.1068","Birinci sınıf usta","Sa",3.0,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",6.0,230.0],["10.100.1062","Düz işçi","Sa",6.0,205.0]]],
[900,"19.100.1087","Alüminyum doğrama imalat atelyesinin 1 saatli ücreti","Sa",4634.86,4634.86,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",1.0,445.0],["10.100.1068","Birinci sınıf usta","Sa",3.0,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",6.0,230.0],["10.100.1062","Düz işçi","Sa",6.0,205.0],["10.100.1059","Yağcı","Sa",0.1,230.0],["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",1.0,445.0]]],
[901,"19.100.1088","Plastik doğrama imalat atelyesinin 1 saatli ücreti","Sa",4603.42,4603.42,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",1.0,445.0],["10.100.1068","Birinci sınıf usta","Sa",3.0,310.0],["10.100.1069","Birinci sınıf usta yardımcısı","Sa",6.0,230.0],["10.100.1062","Düz işçi","Sa",6.0,205.0],["10.100.1059","Yağcı","Sa",0.1,230.0],["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",1.0,445.0]]],
[902,"19.100.1089","Demir doğrama imalat atelyesinin 1 saatli ücreti","Sa",5235.8,5235.8,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",1.0,445.0],["10.100.1018","Sıcak demirci ustası","Sa",3.5,310.0],["10.100.1046","Sıcak demirci usta yardımcısı","Sa",6.5,230.0],["10.100.1062","Düz işçi","Sa",6.5,205.0],["10.100.1059","Yağcı","Sa",0.2,230.0],["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",1.0,445.0]]],
[903,"19.100.1090","Tünel kalıp imalat atelyesinin 1 saatlik ücreti","Sa",5281.3,5281.3,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",1.0,445.0],["10.100.1018","Sıcak demirci ustası","Sa",3.5,310.0],["10.100.1046","Sıcak demirci usta yardımcısı","Sa",6.5,230.0],["10.100.1023","Boyacı ustası","Sa",0.1,310.0],["10.100.1062","Düz işçi","Sa",6.5,205.0],["10.100.1059","Yağcı","Sa",0.2,230.0],["10.130.9991","Su","m³",0.5,55.0]]],
[904,"19.100.1091","Ahşap doğrama imalat atelyesinin 1 saatlik ücreti","Sa",6059.7,6059.7,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",1.0,445.0],["10.100.1009","Marangoz ustası","Sa",6.0,310.0],["10.100.1008","Doğramacı ustası","Sa",3.0,310.0],["10.100.1017","Dülger ustası","Sa",1.0,310.0],["10.100.1041","Marangoz usta yardımcısı","Sa",1.0,230.0],["10.100.1062","Düz işçi","Sa",6.2,205.0],["19.100.1112","Forklift","Sa",0.1,885.98]]],
[905,"19.100.1095","Eksiz Oluk Makinasının 1 saatlik ücreti","Sa",465.39,465.39,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0]]],
[906,"19.100.1096","(Ekskavatör (280 hp) + Tek Depolu Hareketli Kireç Silosu (130 hp) + Karıştırıcı Uç +","Sa",15916.26,15916.26,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.4,445.0],["10.100.1055","Operatör makinist","Sa",2.4,355.0],["10.100.1059","Yağcı","Sa",0.8,230.0]]],
[907,"19.100.1097","(Ekskavatör (280 hp) + Çift Depolu Hareketli Kireç Silosu (130 hp) + Karıştırıcı Uç +","Sa",17797.26,17797.26,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.4,445.0],["10.100.1055","Operatör makinist","Sa",2.4,355.0],["10.100.1059","Yağcı","Sa",0.8,230.0]]],
[908,"19.100.1099","Zemin Stabilizasyon Makinalarının (Kireç Serme Makinası - 250 hp) bir saatlik ücreti","Sa",5569.33,5569.33,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.4,445.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",0.4,230.0],["10.100.1054","Makinist","Sa",1.0,310.0]]],
[909,"19.100.1100","Sıva makinası'nın 1 saatlik ücreti","Sa",749.85,749.85,"Makine/İşçilik","",[["10.100.1054","Makinist","Sa",1.0,310.0],["10.100.1037","Beton Pompa Operatörü","Sa",1.2,375.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0]]],
[910,"19.100.1101","Mobil Beton Pompasının Bir Saatlik Ücreti (420 HP)","Sa",6653.59,6653.59,"Makine/İşçilik","",[["10.100.1037","Beton Pompa Operatörü","Sa",1.2,375.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0]]],
[911,"19.100.1102","Vinç'in 1 saatlik ücreti","Sa",4108.04,4108.04,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0]]],
[912,"19.100.1103","60 Ton kapasiteli mobil vinç'in 1 saatlik ücreti","Sa",5309.19,5309.19,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.44,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.48,445.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0]]],
[913,"19.100.1104","Mobil vinç'in (60ton, 240 HP) 1 saatlik ücreti","Sa",4850.26,4850.26,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1085","Kule vinç operatörü","Sa",1.2,485.0]]],
[914,"19.100.1105","Kule vinçin 1 saatlik ücreti","Sa",3590.32,3590.32,"Makine/İşçilik","",[["10.100.1085","Kule vinç operatörü","Sa",1.2,485.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0]]],
[915,"19.100.1106","Paletli delgi makinası'nın 1 saatlik ücreti","Sa",3393.51,3393.51,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1055","Operatör makinist","Sa",2.4,355.0]]],
[916,"19.100.1107","Jet grouting ekipmanı ile delgi makinası'nın 1 saatlik ücreti","Sa",9662.04,9662.04,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",2.4,355.0],["10.100.1063","Erbab işçi","Sa",1.2,250.0],["10.100.1062","Düz işçi","Sa",1.2,205.0]]],
[917,"19.100.1108","Mikro tünel sistemi ile boru sürme makinası'nın 1 saatlik ücreti","Sa",11814.11,11814.11,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.72,445.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.72,445.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0]]],
[918,"19.100.1109","Mikro tünel sistemi ile boru sürme makinası'nın 1 saatlik ücreti","Sa",43083.33,43083.33,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",0.72,445.0],["10.100.1054","Makinist","Sa",1.44,310.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1081","Elektrik ustası","Sa",1.0,310.0]]],
[919,"19.100.1110","Matkap'ın 1 saatlik ücreti","Sa",369.82,369.82,"Makine/İşçilik","",[["10.100.1081","Elektrik ustası","Sa",1.0,310.0]]],
[920,"19.100.1111","Demir kesme ve bükme makinası'nın 1 saatlik ücreti","Sa",50.47,50.47,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0]]],
[921,"19.100.1112","Forklift'in 1 saatlik ücreti","Sa",885.98,1111.93,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0]]],
[922,"19.100.1113","Mobil vinç'in 1 saatlik ücreti","Sa",1402.45,1816.56,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0]]],
[923,"19.100.1114","Çift kompenantlı yalıtım malzemesi dozaj miks makinasının bir saatlik ücreti","Sa",930.4,930.4,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.4,445.0]]],
[924,"19.100.1115","Grab makinesi'nin 1 saatlik ücreti","Sa",15094.73,15094.73,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.4,445.0]]],
[925,"19.100.1116","Hidrofreze makinesi'nin 1 saatlik ücreti","Sa",59196.71,59196.71,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1057","Operatör yardımcısı","Sa",1.2,300.0],["10.100.1059","Yağcı","Sa",1.0,230.0],["10.100.1060","Formen","Sa",0.4,445.0],["10.100.1062","Düz işçi","Sa",2.0,205.0],["10.100.1021","Kaynakçı ustası","Sa",0.25,310.0]]],
[926,"19.100.1117","Bentonit ünitesi ve Desander 1 saatlik ücreti (Grab için)","Sa",17460.28,17460.28,"Makine/İşçilik","",[["19.100.1041","Motopomp'un (45 Ps.) 1 Saatlik Ücreti","Sa",4.0,874.78],["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1057","Operatör yardımcısı","Sa",1.2,300.0]]],
[927,"19.100.1118","Bentonit ünitesi ve Desander 1 saatlik ücreti (Hidrofreze için)","Sa",23441.37,23441.37,"Makine/İşçilik","",[["19.100.1119","Salyangoz pompanın 1 saatlik ücreti","Sa",1.0,2069.15],["19.100.1041","Motopomp'un (45 Ps.) 1 Saatlik Ücreti","Sa",8.0,874.78],["10.100.1055","Operatör makinist","Sa",1.2,355.0],["10.100.1057","Operatör yardımcısı","Sa",1.2,300.0]]],
[928,"19.100.1119","Salyangoz pompanın 1 saatlik ücreti","Sa",2069.15,2069.15,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.5,55.0],["10.100.1060","Formen","Sa",1.0,445.0],["10.100.1018","Sıcak demirci ustası","Sa",3.0,310.0],["10.100.1046","Sıcak demirci usta yardımcısı","Sa",3.0,230.0],["10.100.1062","Düz işçi","Sa",3.0,205.0],["10.100.1059","Yağcı","Sa",0.2,230.0],["10.100.1023","Boyacı ustası","Sa",1.0,310.0]]],
[929,"19.100.1122","Bor katkılı selüloz yünü püskürtme makinası'nın 1 saatlik ücreti","Sa",472.75,472.75,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.15,55.0],["10.100.1055","Operatör makinist","Sa",1.0,355.0],["10.100.1062","Düz işçi","Sa",0.65,205.0]]],
[930,"19.100.2001","El arabası ile her cins malzeme ve kayadan başka kazının taşınması (50 metre'ye)","ton",133.25,133.25,"Makine/İşçilik","",[["10.100.1062","Düz işçi","Sa",0.65,205.0],["10.100.1062","Düz işçi","Sa",0.78,205.0]]],
[931,"19.100.2002","El arabası ile her cins malzeme ve kayadan başka kazının taşınması (60 metre'ye)","ton",159.9,159.9,"Makine/İşçilik","",[["10.100.1062","Düz işçi","Sa",0.78,205.0],["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.014,2183.46],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.0083,1686.26],["19.100.1035","Elek","Sa",0.025,1334.95],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.0083,1686.26],["10.100.1062","Düz işçi","Sa",0.5,205.0],["19.100.1037","Motopomp (5 Ps)","Sa",0.1,305.51],["10.100.1060","Formen","Sa",0.25,445.0]]],
[932,"19.100.2004","Makina ile elenmiş, yıkanmış granülometrik kum-çakıl hazırlanması","m³",336.24,336.24,"Makine/İşçilik","",[["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.014,2183.46],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.0083,1686.26],["19.100.1035","Elek","Sa",0.025,1334.95],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.0083,1686.26],["10.100.1062","Düz işçi","Sa",0.5,205.0],["19.100.1037","Motopomp (5 Ps)","Sa",0.1,305.51],["10.100.1060","Formen","Sa",0.25,445.0],["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.014,2183.46]]],
[933,"19.100.2003","Makina ile kum çakıl hazırlanması","m³",133.07,133.07,"Makine/İşçilik","",[["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.014,2183.46],["10.100.1062","Düz işçi","Sa",0.5,205.0],["19.100.1006","Ekskavatör (paletli) (210-259 HP)","Sa",0.014,2971.46],["19.100.1029","Yükleyici (100 HP)","Sa",0.014,1510.89],["19.100.1036","Eleme makinası","Sa",0.025,1254.75],["19.100.1038","Motopomp","Sa",0.1,364.62],["10.100.1060","Formen","Sa",0.25,445.0],["19.100.1029","Yükleyici (100 HP)","Sa",0.025,1510.89]]],
[934,"19.100.2005","Makina ile elenmiş, yıkanmış granülometrik kum-çakıl hazırlanması","m³",279.6,279.6,"Makine/İşçilik","",[["19.100.1006","Ekskavatör (paletli) (210-259 HP)","Sa",0.014,2971.46],["19.100.1029","Yükleyici (100 HP)","Sa",0.014,1510.89],["19.100.1036","Eleme makinası","Sa",0.025,1254.75],["19.100.1038","Motopomp","Sa",0.1,364.62],["10.100.1060","Formen","Sa",0.25,445.0],["19.100.1029","Yükleyici (100 HP)","Sa",0.025,1510.89],["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.014,2183.46],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.0083,1686.26]]],
[935,"19.100.2006","Makina ile elenmiş, yıkanmış iki tane sınıfı ayrılmış granülometrik kum-çakıl hazırlanması","m³",380.74,380.74,"Makine/İşçilik","",[["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.014,2183.46],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.0083,1686.26],["19.100.1035","Elek","Sa",0.025,1334.95],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.0083,1686.26],["10.100.1062","Düz işçi","Sa",0.5,205.0],["19.100.1037","Motopomp (5 Ps)","Sa",0.1,305.51],["10.100.1060","Formen","Sa",0.35,445.0],["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.014,2183.46]]],
[936,"19.100.2007","Makina ile elenmiş, yıkanmış ince sıva veya derz kumu hazırlanması","m³",429.99,429.99,"Makine/İşçilik","",[["19.100.1001","Ekskavatör (100-139 HP)","Sa",0.014,2183.46],["19.100.1037","Motopomp (5 Ps)","Sa",0.1,305.51],["10.100.1062","Düz işçi","Sa",1.0,205.0],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.0083,1686.26],["19.100.1035","Elek","Sa",0.025,1334.95],["19.100.1028","Lastik tekerlekli yükleyici (80 HP)","Sa",0.0083,1686.26],["10.100.1062","Düz işçi","Sa",0.5,205.0],["19.100.1006","Ekskavatör (paletli) (210-259 HP)","Sa",0.014,2971.46]]],
[937,"19.100.2008","Makina ile elenmiş, yıkanmış ince sıva veya derz kumu hazırlanması","m³",281.67,281.67,"Makine/İşçilik","",[["19.100.1006","Ekskavatör (paletli) (210-259 HP)","Sa",0.014,2971.46],["19.100.1029","Yükleyici (100 HP)","Sa",0.014,1510.89],["19.100.1036","Eleme makinası","Sa",0.025,1254.75],["19.100.1038","Motopomp","Sa",0.1,364.62],["19.100.1029","Yükleyici (100 HP)","Sa",0.1,1510.89]]],
[938,"19.100.2014","Vidye kron hazırlanması","adet",5758.0,5758.0,"Makine/İşçilik","",[["10.100.1060","Formen","Sa",4.4,445.0],["10.100.1063","Erbab işçi","Sa",4.4,250.0],["19.100.1106","Paletli delgi makinası (160 HP)","Sa",0.011,3393.51],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.011,310.0],["10.100.1063","Erbab işçi","Sa",0.044,250.0],["19.100.1008","Ekskavatör (paletli) (260-299 HP)","Sa",0.05,3421.44]]],
[939,"19.100.2015","Ocakta taş hazırlanması","m³",345.39,345.39,"Makine/İşçilik","",[["19.100.1106","Paletli delgi makinası (160 HP)","Sa",0.011,3393.51],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.011,310.0],["10.100.1063","Erbab işçi","Sa",0.044,250.0],["19.100.1008","Ekskavatör (paletli) (260-299 HP)","Sa",0.05,3421.44],["19.100.1029","Yükleyici (100 HP)","Sa",0.03,1510.89],["10.100.1062","Düz işçi","Sa",0.25,205.0],["19.100.2015","Ocakta hazırlanan taş","m³",1.0,345.39],["19.100.1029","Yükleyici (100 HP)","Sa",0.02,1510.89]]],
[940,"19.100.2016","Taştan konkasörle kırılmış ve elenmiş 70 mm'ye kadar kırma taş hazırlanması","m³",421.95,421.95,"Makine/İşçilik","",[["19.100.2015","Ocakta hazırlanan taş","m³",1.0,345.39],["19.100.1029","Yükleyici (100 HP)","Sa",0.02,1510.89],["19.100.1034","Konkasör","Sa",0.008,4659.08],["19.100.1029","Yükleyici (100 HP)","Sa",0.006,1510.89],["19.100.2015","Ocakta hazırlanan taş","m³",1.0,345.39],["19.100.1029","Yükleyici (100 HP)","Sa",0.02,1510.89],["19.100.1034","Konkasör","Sa",0.011,4659.08]]],
[941,"19.100.2017","Taştan konkasörle kırılmış ve elenmiş 30 mm'ye kadar kırma taş hazırlanması","m³",480.43,480.43,"Makine/İşçilik","",[["19.100.2015","Ocakta hazırlanan taş","m³",1.0,345.39],["19.100.1029","Yükleyici (100 HP)","Sa",0.02,1510.89],["19.100.1034","Konkasör","Sa",0.011,4659.08],["10.100.1060","Formen","Sa",0.1,445.0],["19.100.1029","Yükleyici (100 HP)","Sa",0.006,1510.89],["10.100.1063","Erbab işçi","Sa",2.0,250.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[942,"19.100.2018","Kazıdan taş hazırlanması","m³",602.5,602.5,"Makine/İşçilik","",[["10.100.1063","Erbab işçi","Sa",2.0,250.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1063","Erbab işçi","Sa",4.0,250.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[943,"19.100.2019","Toplama suretiyle taş hazırlanması","m³",1102.5,1102.5,"Makine/İşçilik","",[["10.100.1063","Erbab işçi","Sa",4.0,250.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["19.100.1106","Paletli delgi makinası (160 HP)","Sa",0.011,3393.51],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.011,310.0],["10.100.1063","Erbab işçi","Sa",0.044,250.0]]],
[944,"19.100.2020","Ocakta çaplanmış moloz taşı hazırlanması","m³",1400.39,1400.39,"Makine/İşçilik","",[["19.100.1106","Paletli delgi makinası (160 HP)","Sa",0.011,3393.51],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.011,310.0],["10.100.1063","Erbab işçi","Sa",0.044,250.0],["19.100.1008","Ekskavatör (paletli) (260-299 HP)","Sa",0.05,3421.44],["19.100.1029","Yükleyici (100 HP)","Sa",0.03,1510.89],["10.100.1062","Düz işçi","Sa",0.25,205.0],["10.100.1001","Taşçı ustası","Sa",3.0,310.0],["10.100.1063","Erbab işçi","Sa",0.5,250.0]]],
[945,"19.100.2021","Ocakta, kemer inşaatı için çaplanmış moloz taşı hazırlanması","m³",6101.4,6101.4,"Makine/İşçilik","",[["10.160.1005","Kapsül","adet",1.5,19.75],["10.160.1004","Fitil","m",1.5,12.0],["10.100.1062","Düz işçi","Sa",0.2,205.0],["19.100.1023","Kompresör","Sa",0.145,2508.08],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.25,310.0],["10.100.1063","Erbab işçi","Sa",2.5,250.0],["10.100.1063","Erbab işçi","Sa",0.35,250.0],["10.100.1001","Taşçı ustası","Sa",8.0,310.0]]],
[946,"19.100.2022","Kazıdan çıkan taşlardan çaplanmış moloz taşı hazırlanması","m³",2872.5,2872.5,"Makine/İşçilik","",[["10.100.1063","Erbab işçi","Sa",2.0,250.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1001","Taşçı ustası","Sa",6.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1063","Erbab işçi","Sa",2.0,250.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[947,"19.100.2023","Kazıdan çıkan taşlardan kemer inşaat için çaplanmış moloz taşı hazırlanması","m³",3492.5,3492.5,"Makine/İşçilik","",[["10.100.1063","Erbab işçi","Sa",2.0,250.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1001","Taşçı ustası","Sa",8.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.160.1005","Kapsül","adet",2.0,19.75],["10.160.1004","Fitil","m",2.0,12.0],["10.100.1062","Düz işçi","Sa",0.3,205.0]]],
[948,"19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",4240.62,4240.62,"Makine/İşçilik","",[["10.160.1005","Kapsül","adet",2.0,19.75],["10.160.1004","Fitil","m",2.0,12.0],["10.100.1062","Düz işçi","Sa",0.3,205.0],["19.100.1023","Kompresör","Sa",0.2,2508.08],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.25,310.0],["10.100.1063","Erbab işçi","Sa",2.5,250.0],["10.100.1063","Erbab işçi","Sa",0.4,250.0],["10.100.1001","Taşçı ustası","Sa",8.0,310.0]]],
[949,"19.100.2025","Ocakta özel yonu taşı taslağı hazırlanması","m³",5490.87,5490.87,"Makine/İşçilik","",[["10.160.1005","Kapsül","adet",2.0,19.75],["10.160.1004","Fitil","m",2.0,12.0],["10.100.1062","Düz işçi","Sa",0.35,205.0],["19.100.1023","Kompresör","Sa",0.2,2508.08],["10.100.1011","Lağımcı (Ateşleme ustası)","Sa",0.25,310.0],["10.100.1063","Erbab işçi","Sa",2.5,250.0],["10.100.1063","Erbab işçi","Sa",0.4,250.0],["10.100.1001","Taşçı ustası","Sa",12.0,310.0]]],
[950,"19.100.2026","Ocakta hazırlanmış kaba yonu taşı taslağından kaba yonu taşı hazırlanması","m³",8612.81,8612.81,"Makine/İşçilik","",[["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",1.0,4240.62],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",0.3,4240.62],["10.100.1001","Taşçı ustası","Sa",10.0,310.0],["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",1.0,3492.5],["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",0.3,3492.5],["10.100.1001","Taşçı ustası","Sa",10.0,310.0]]],
[951,"19.100.2027","Kazı taşından hazırlanmış yonu taşı taslağından kaba yonu taşı hazırlanması","m³",7640.25,7640.25,"Makine/İşçilik","",[["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",1.0,3492.5],["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",0.3,3492.5],["10.100.1001","Taşçı ustası","Sa",10.0,310.0],["19.100.2025","Ocakta özel yonu taşı taslağı hazırlanması","m³",1.0,5490.87],["19.100.2025","Ocakta özel yonu taşı taslağı hazırlanması","m³",0.4,5490.87],["10.100.1001","Taşçı ustası","Sa",13.0,310.0]]],
[952,"19.100.2028","Ocakta hazırlanmış özel yonu taşı taslağından kaba yonu taşı hazırlanması","m³",11717.22,11717.22,"Makine/İşçilik","",[["19.100.2025","Ocakta özel yonu taşı taslağı hazırlanması","m³",1.0,5490.87],["19.100.2025","Ocakta özel yonu taşı taslağı hazırlanması","m³",0.4,5490.87],["10.100.1001","Taşçı ustası","Sa",13.0,310.0],["19.100.2040","","m³",1.0,4732.5],["19.100.2040","","m³",0.4,4732.5]]],
[953,"19.100.2029","Kazı taşından hazırlanmış yonu taşı taslağından kaba yonu taşı hazırlanması","m³",10655.5,10655.5,"Makine/İşçilik","",[["19.100.2040","","m³",1.0,4732.5],["19.100.2040","","m³",0.4,4732.5],["10.100.1001","Taşçı ustası","Sa",13.0,310.0],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",1.0,4240.62],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",0.4,4240.62],["10.100.1001","Taşçı ustası","Sa",25.0,310.0]]],
[954,"19.100.2030","Ocakta hazırlanmış yonu taşı taslağından ince yonu taşı hazırlanması","m³",13686.87,13686.87,"Makine/İşçilik","",[["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",1.0,4240.62],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",0.4,4240.62],["10.100.1001","Taşçı ustası","Sa",25.0,310.0],["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",1.0,3492.5],["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",0.4,3492.5],["10.100.1001","Taşçı ustası","Sa",25.0,310.0]]],
[955,"19.100.2031","Kazı taşından hazırlanmış yonu taşı taslağından ince yonu taşı hazırlanması","m³",12639.5,12639.5,"Makine/İşçilik","",[["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",1.0,3492.5],["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",0.4,3492.5],["10.100.1001","Taşçı ustası","Sa",25.0,310.0],["19.100.2025","Ocakta özel yonu taşı taslağı hazırlanması","m³",1.0,5490.87],["19.100.2025","Ocakta özel yonu taşı taslağı hazırlanması","m³",0.4,5490.87],["10.100.1001","Taşçı ustası","Sa",35.0,310.0]]],
[956,"19.100.2032","Ocakta hazırlanmış özel yonu taşı taslağından özel ince yonu taşı hazırlanması","m³",18537.22,18537.22,"Makine/İşçilik","",[["19.100.2025","Ocakta özel yonu taşı taslağı hazırlanması","m³",1.0,5490.87],["19.100.2025","Ocakta özel yonu taşı taslağı hazırlanması","m³",0.4,5490.87],["10.100.1001","Taşçı ustası","Sa",35.0,310.0],["19.100.2040","","m³",1.0,4732.5],["19.100.2040","","m³",0.4,4732.5]]],
[957,"19.100.2033","Kazı taşından hazırlanmış özel yonu taşı taslağından özel ince yonu taşı hazırlanması","m³",17475.5,17475.5,"Makine/İşçilik","",[["19.100.2040","","m³",1.0,4732.5],["19.100.2040","","m³",0.4,4732.5],["10.100.1001","Taşçı ustası","Sa",35.0,310.0],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",1.0,4240.62],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",0.6,4240.62],["10.100.1001","Taşçı ustası","Sa",60.0,310.0]]],
[958,"19.100.2034","Ocakta hazırlanmış yonu taşı taslağından kesme taş hazırlanması","m³",25384.99,25384.99,"Makine/İşçilik","",[["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",1.0,4240.62],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",0.6,4240.62],["10.100.1001","Taşçı ustası","Sa",60.0,310.0],["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",1.0,3492.5],["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",0.6,3492.5],["10.100.1001","Taşçı ustası","Sa",60.0,310.0]]],
[959,"19.100.2035","Kazı taşından hazırlanmış yonu taşı taslağından kesme taş hazırlanması","m³",24188.0,24188.0,"Makine/İşçilik","",[["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",1.0,3492.5],["19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",0.6,3492.5],["10.100.1001","Taşçı ustası","Sa",60.0,310.0],["10.100.1001","Taşçı ustası","Sa",10.0,310.0],["10.100.1001","Taşçı ustası","Sa",15.0,310.0]]],
[960,"19.100.2036","Yumuşak kesme taş hazırlanması","m³",7750.0,7750.0,"Makine/İşçilik","",[["10.100.1001","Taşçı ustası","Sa",10.0,310.0],["10.100.1001","Taşçı ustası","Sa",15.0,310.0],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",1.0,4240.62],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",0.6,4240.62],["10.100.1001","Taşçı ustası","Sa",60.0,310.0]]],
[961,"19.100.2037","Doğal taşlarla, alt yüzü kabaca murçlanmış ince yonu kaplama taşı hazırlanması","m³",25384.99,25384.99,"Makine/İşçilik","",[["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",1.0,4240.62],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",0.6,4240.62],["10.100.1001","Taşçı ustası","Sa",60.0,310.0],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",1.0,4240.62],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",0.6,4240.62],["10.100.1001","Taşçı ustası","Sa",70.0,310.0]]],
[962,"19.100.2038","Doğal taşlarla, alt yüzü kabaca tesviye edilmiş ince yonu kaplama taşı hazırlanması","m³",28484.99,28484.99,"Makine/İşçilik","",[["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",1.0,4240.62],["19.100.2024","Ocakta yonu taşı taslağı hazırlanması","m³",0.6,4240.62],["10.100.1001","Taşçı ustası","Sa",70.0,310.0],["10.100.1063","Erbab işçi","Sa",2.0,250.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1001","Taşçı ustası","Sa",8.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[963,"19.100.2039","Kazı taşından yonu taşı taslağı hazırlanması","m³",3492.5,3492.5,"Makine/İşçilik","",[["10.100.1063","Erbab işçi","Sa",2.0,250.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1001","Taşçı ustası","Sa",8.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0],["10.100.1063","Erbab işçi","Sa",2.0,250.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[964,"19.100.2040","Kazı taşından çıkan taşlardan özel yonu taşı taslağı hazırlanması","m³",4732.5,4732.5,"Makine/İşçilik","",[["10.100.1063","Erbab işçi","Sa",2.0,250.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1001","Taşçı ustası","Sa",12.0,310.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",0.5,205.0]]],
[965,"19.100.2041","Ocak taşından (0-0.005 ton kategorideki) taşın hazırlanması","ton",71.84,71.84,"Makine/İşçilik","",[]],
[966,"19.100.2042","Kazı taşından (0-0.005 ton kategorideki) taşın hazırlanması","ton",22.86,22.86,"Makine/İşçilik","",[]],
[967,"19.100.2043","Ocak taşından (0.005-0.100 ton kategorideki)taşın hazırlanması","ton",149.01,149.01,"Makine/İşçilik","",[]],
[968,"19.100.2044","Kazı taşından (0.005-0.100 ton kategorideki) taşın hazırlanması","ton",38.06,38.06,"Makine/İşçilik","",[]],
[969,"19.100.2045","Ocak taşından (0-0.250 ton kategorideki)taşın hazırlanması","ton",164.82,164.82,"Makine/İşçilik","",[]],
[970,"19.100.2046","Kazı taşından (0-0.250 ton kategorideki) taşın hazırlanması","ton",41.88,41.88,"Makine/İşçilik","",[]],
[971,"19.100.2047","Ocak taşından (0.100-0.250 ton kategorideki)taşın hazırlanması","ton",170.87,170.87,"Makine/İşçilik","",[]],
[972,"19.100.2048","Kazı taşından (0.100-0.250 ton kategorideki) taşın hazırlanması","ton",47.57,47.57,"Makine/İşçilik","",[["10.160.1004","Fitil","m",0.85,12.0],["10.160.1005","Kapsül","adet",0.85,19.75],["19.100.1023","Kompresör","Sa",0.095,2508.08],["19.100.1005","Ekskavatör ((210-259 HP)","Sa",0.012,3374.35]]],
[973,"19.100.2049","Ocak taşından (0-0.400 ton kategorideki)taşın hazırlanması","m³",315.95,315.95,"Makine/İşçilik","",[["10.160.1004","Fitil","m",0.85,12.0],["10.160.1005","Kapsül","adet",0.85,19.75],["19.100.1023","Kompresör","Sa",0.095,2508.08],["19.100.1005","Ekskavatör ((210-259 HP)","Sa",0.012,3374.35]]],
[974,"19.100.2050","Kazı taşından (0-0.400 ton kategorideki) taşın hazırlanması","ton",51.4,51.4,"Makine/İşçilik","",[]],
[975,"19.100.2051","Ocak taşından (0.250-0.400 ton kategorideki)taşın hazırlanması","ton",182.07,182.07,"Makine/İşçilik","",[]],
[976,"19.100.2052","Ocak taşından (0.4-2 ton kategorideki) taşın hazırlanması","ton",188.13,188.13,"Makine/İşçilik","",[]],
[977,"19.100.2053","Kazı taşından (0.4-2.0 ton kategorideki) taşın hazırlanması","ton",64.75,64.75,"Makine/İşçilik","",[]],
[978,"19.100.2054","Ocak taşından (2-6 ton kategorideki)taşın hazırlanması","ton",234.75,234.75,"Makine/İşçilik","",[]],
[979,"19.100.2055","Kazı taşından (2-6 ton kategorideki) taşın hazırlanması","ton",85.9,85.9,"Makine/İşçilik","",[]],
[980,"19.100.2056","Ocak taşından (6-15 ton kategorideki) taşın hazırlanması","ton",295.29,295.29,"Makine/İşçilik","",[]],
[981,"19.100.2057","Kazı taşından (6-15 ton kategorideki) taşın hazırlanması","ton",105.06,105.06,"Makine/İşçilik","",[]],
[982,"19.100.2058","Ocak taşından (15ton'dan yukarı kategorideki) taşın hazırlanması","ton",348.87,348.87,"Makine/İşçilik","",[]],
[983,"19.100.2059","Kazı taşından (15ton'dan yukarı kategorideki) taşın hazırlanması","ton",134.15,134.15,"Makine/İşçilik","",[["19.100.1019","Traktör Buldozer (185 HP)","Sa",0.005,3141.71],["19.100.1005","Ekskavatör ((210-259 HP)","Sa",0.005,3374.35]]],
[984,"19.100.2060","Ocak artığının ocakta toplanması","m³",32.58,32.58,"Makine/İşçilik","",[["19.100.1019","Traktör Buldozer (185 HP)","Sa",0.005,3141.71],["19.100.1005","Ekskavatör ((210-259 HP)","Sa",0.005,3374.35],["19.100.2049","","m³",0.4,315.95],["19.100.1019","Traktör Buldozer (185 HP)","Sa",0.005,3141.71]]],
[985,"19.100.2061","Ocakta kalan kategori fazlası taş","m³",142.09,142.09,"Makine/İşçilik","",[["19.100.2049","","m³",0.4,315.95],["19.100.1019","Traktör Buldozer (185 HP)","Sa",0.005,3141.71],["10.130.9991","Su","m³",1.0,55.0],["10.100.1062","Düz işçi","Sa",3.0,205.0]]],
[986,"19.100.2062","İş yerinde 1 m³ söndürülmemiş parça kalker kireci hazırlanması","m³",2215.0,2215.0,"Makine/İşçilik","",[["10.130.9991","Su","m³",1.0,55.0],["10.100.1062","Düz işçi","Sa",3.0,205.0],["10.100.1062","Düz işçi","Sa",1.5,205.0],["19.100.1037","Motopomp (5 Ps)","Sa",0.067,305.51]]],
[987,"19.100.2063","Elle 1 m³ su hazırlanması","m³",307.5,307.5,"Makine/İşçilik","",[["10.100.1062","Düz işçi","Sa",1.5,205.0],["19.100.1037","Motopomp (5 Ps)","Sa",0.067,305.51],["10.160.1035","Elektrot","adet",23.0,1.92],["19.100.1061","Kaynak Makinası (25 HP)","Sa",1.1,492.91],["10.100.1021","Kaynakçı ustası","Sa",1.1,310.0],["10.100.1062","Düz işçi","Sa",1.1,205.0]]],
[988,"19.100.2064","Motorlu tulumba ile su hazırlanması","m³",20.47,20.47,"Makine/İşçilik","",[["19.100.1037","Motopomp (5 Ps)","Sa",0.067,305.51],["10.160.1035","Elektrot","adet",23.0,1.92],["19.100.1061","Kaynak Makinası (25 HP)","Sa",1.1,492.91],["10.100.1021","Kaynakçı ustası","Sa",1.1,310.0],["10.100.1062","Düz işçi","Sa",1.1,205.0],["10.130.9991","Su","m³",0.2,55.0],["10.100.1062","Düz işçi","Sa",2.3,205.0],["10.100.1062","Düz işçi","Sa",2.5,205.0]]],
[989,"19.100.2065","Çelik boru başlarının kaynakla bağlanması","m",1152.86,1152.86,"Makine/İşçilik","",[["10.160.1035","Elektrot","adet",23.0,1.92],["19.100.1061","Kaynak Makinası (25 HP)","Sa",1.1,492.91],["10.100.1021","Kaynakçı ustası","Sa",1.1,310.0],["10.100.1062","Düz işçi","Sa",1.1,205.0],["10.130.9991","Su","m³",0.2,55.0],["10.100.1062","Düz işçi","Sa",2.3,205.0],["10.100.1062","Düz işçi","Sa",2.5,205.0]]],
[990,"19.100.2416","Çamur harcı hazırlanması","m³",1070.0,1070.0,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.2,55.0],["10.100.1062","Düz işçi","Sa",2.3,205.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.240.3201","Mermer pirinci (beyaz)","ton",1.45,420.0],["10.130.1235","Beyaz Portland Kalkerli Çimento","ton",0.65,5100.0],["10.130.9991","Su","m³",0.3,55.0],["10.100.1062","Düz işçi","Sa",3.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[991,"19.100.2426","Mozayik harcı yapılması (beyaz çimentolu)","m³",4863.0,4863.0,"Makine/İşçilik","",[["10.240.3201","Mermer pirinci (beyaz)","ton",1.45,420.0],["10.130.1235","Beyaz Portland Kalkerli Çimento","ton",0.65,5100.0],["10.130.9991","Su","m³",0.3,55.0],["10.100.1062","Düz işçi","Sa",3.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.9991","Su","m³",0.6,55.0],["10.100.1044","Sıvacı usta yardımcısı","Sa",4.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[992,"19.100.2432","Saten alçı harcı hazırlanması","m³",4243.0,4243.0,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.6,55.0],["10.100.1044","Sıvacı usta yardımcısı","Sa",4.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.9991","Su","m³",0.395,55.0],["10.100.1044","Sıvacı usta yardımcısı","Sa",3.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[993,"19.100.2433","Perlitli alçı harcı hazırlanması","m³",2219.73,2219.73,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.395,55.0],["10.100.1044","Sıvacı usta yardımcısı","Sa",3.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.9991","Su","m³",0.59,55.0],["10.100.1038","Alçı Levha Usta Yardımcısı","Sa",4.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[994,"19.100.2434","Derz dolgu alçısı harcı hazırlanması","m³",4593.95,4593.95,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.59,55.0],["10.100.1038","Alçı Levha Usta Yardımcısı","Sa",4.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.9991","Su","m³",0.59,55.0],["10.100.1038","Alçı Levha Usta Yardımcısı","Sa",4.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[995,"19.100.2435","Yapıştırma alçısı harcı hazırlanması","m³",4593.95,4593.95,"Makine/İşçilik","",[["10.130.9991","Su","m³",0.59,55.0],["10.100.1038","Alçı Levha Usta Yardımcısı","Sa",4.5,230.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.2,2550.0],["10.130.9991","Su","m³",0.13,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0]]],
[996,"19.100.2501","200 Kg çimento dozlu harç yapılması","m³",1779.65,1779.65,"Makine/İşçilik","",[["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.2,2550.0],["10.130.9991","Su","m³",0.13,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.25,2550.0]]],
[997,"19.100.2502","250 Kg çimento dozlu tesviye harcı yapılması","m³",1702.15,1702.15,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.25,2550.0],["10.130.9991","Su","m³",0.13,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.2,2550.0],["10.130.9991","Su","m³",0.2,55.0]]],
[998,"19.100.2503","200 Kg çimento dozlu harç yapılması","m³",1578.5,1578.5,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.2,2550.0],["10.130.9991","Su","m³",0.2,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.25,2550.0],["10.130.9991","Su","m³",0.215,55.0]]],
[999,"19.100.2504","250 kg çimento dozlu harç yapılması (kargir işlerde)","m³",1706.83,1706.83,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.25,2550.0],["10.130.9991","Su","m³",0.215,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.3,2550.0],["10.130.9991","Su","m³",0.23,55.0]]],
[1000,"19.100.2505","300 Kg çimento dozlu harç yapılması","m³",1835.15,1835.15,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.3,2550.0],["10.130.9991","Su","m³",0.23,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1007","","m³",1.0,431.0],["10.130.1209","","ton",0.3,2550.0],["10.130.9991","Su","m³",0.23,55.0]]],
[1001,"19.100.2506","300 Kg çimento dozlu ince harç yapılması","m³",1926.15,1926.15,"Makine/İşçilik","",[["10.130.1007","","m³",1.0,431.0],["10.130.1209","","ton",0.3,2550.0],["10.130.9991","Su","m³",0.23,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.35,2550.0],["10.130.9991","Su","m³",0.245,55.0]]],
[1002,"19.100.2507","350 Kg çimento dozlu harç yapılması","m³",1963.48,1963.48,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.35,2550.0],["10.130.9991","Su","m³",0.245,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1007","","m³",1.0,431.0],["10.130.1209","","ton",0.35,2550.0],["10.130.9991","Su","m³",0.245,55.0]]],
[1003,"19.100.2508","350 Kg çimento dozlu ince harç yapılması","m³",2054.48,2054.48,"Makine/İşçilik","",[["10.130.1007","","m³",1.0,431.0],["10.130.1209","","ton",0.35,2550.0],["10.130.9991","Su","m³",0.245,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.4,2550.0],["10.130.9991","Su","m³",0.26,55.0]]],
[1004,"19.100.2509","400 Kg çimento dozlu harç yapılması","m³",2091.8,2091.8,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.4,2550.0],["10.130.9991","Su","m³",0.26,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",0.3,340.0],["10.130.1009","","m³",0.7,425.0],["10.130.1209","","ton",0.4,2550.0]]],
[1005,"19.100.2510","Kum ve kırmataş ile 400 Kg çimento dozlu harç yapılması","m³",2151.3,2151.3,"Makine/İşçilik","",[["10.130.1005","","m³",0.3,340.0],["10.130.1009","","m³",0.7,425.0],["10.130.1209","","ton",0.4,2550.0],["10.130.9991","Su","m³",0.26,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1007","","m³",1.0,431.0]]],
[1006,"19.100.2511","400 Kg çimento dozlu ince harç yapılması","m³",2182.8,2182.8,"Makine/İşçilik","",[["10.130.1007","","m³",1.0,431.0],["10.130.1209","","ton",0.4,2550.0],["10.130.9991","Su","m³",0.26,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.45,2550.0],["10.130.9991","Su","m³",0.275,55.0]]],
[1007,"19.100.2512","450 Kg çimento dozlu harç yapılması","m³",2220.13,2220.13,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.45,2550.0],["10.130.9991","Su","m³",0.275,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1007","","m³",1.0,431.0],["10.130.1209","","ton",0.45,2550.0],["10.130.9991","Su","m³",0.275,55.0]]],
[1008,"19.100.2513","450 Kg çimento dozlu ince harç yapılması","m³",2311.13,2311.13,"Makine/İşçilik","",[["10.130.1007","","m³",1.0,431.0],["10.130.1209","","ton",0.45,2550.0],["10.130.9991","Su","m³",0.275,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.5,2550.0],["10.130.9991","Su","m³",0.29,55.0]]],
[1009,"19.100.2514","500 Kg çimento dozlu harç yapılması","m³",2348.45,2348.45,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.5,2550.0],["10.130.9991","Su","m³",0.29,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1007","","m³",1.0,431.0],["10.130.9991","Su","m³",0.29,55.0],["10.130.1209","","ton",0.5,2550.0]]],
[1010,"19.100.2515","500 Kg çimento dozlu ince harç yapılması","m³",2644.45,2644.45,"Makine/İşçilik","",[["10.130.1007","","m³",1.0,431.0],["10.130.9991","Su","m³",0.29,55.0],["10.130.1209","","ton",0.5,2550.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1209","","ton",0.6,2550.0],["10.130.9991","Su","m³",1.0,55.0]]],
[1011,"19.100.2517","600 dozlu çimento şerbeti","m³",1995.0,1995.0,"Makine/İşçilik","",[["10.130.1209","","ton",0.6,2550.0],["10.130.9991","Su","m³",1.0,55.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1209","","ton",0.5,2550.0],["10.130.1007","","m³",0.1,431.0],["10.130.9991","Su","m³",0.9,55.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[1012,"19.100.2518","500 dozlu çimento şerbeti","m³",1777.6,1777.6,"Makine/İşçilik","",[["10.130.1209","","ton",0.5,2550.0],["10.130.1007","","m³",0.1,431.0],["10.130.9991","Su","m³",0.9,55.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.25,2550.0],["10.130.6001","Sönmüş kireç CL 70S","ton",0.076,3059.0]]],
[1013,"19.100.2519","Kireç harcı yapılması (sönmüş kireç torbalı)","m³",1941.51,1941.51,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.25,2550.0],["10.130.6001","Sönmüş kireç CL 70S","ton",0.076,3059.0],["10.130.9991","Su","m³",0.255,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.25,2550.0]]],
[1014,"19.100.2520","Kireç çimento karışımı harç yapılması","m³",1941.51,1941.51,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.25,2550.0],["10.130.9991","Su","m³",0.255,55.0],["10.130.6001","Sönmüş kireç CL 70S","ton",0.076,3059.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1007","","m³",1.0,431.0],["10.130.1209","","ton",0.25,2550.0]]],
[1015,"19.100.2521","0.100 m³/250 Kg kireç-çimento karışımı ince harç yapılması(sönmüş kireç torbalı)","m³",2032.51,2032.51,"Makine/İşçilik","",[["10.130.1007","","m³",1.0,431.0],["10.130.1209","","ton",0.25,2550.0],["10.130.6001","Sönmüş kireç CL 70S","ton",0.076,3059.0],["10.130.9991","Su","m³",0.255,55.0],["10.100.1062","Düz işçi","Sa",2.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.2,2550.0]]],
[1016,"19.100.2522","0.170 m³/200 Kg kireç ve çimento karışımı kaba harç yapılması (sönmüş kireç torbalı)","m³",2077.5,2077.5,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.2,2550.0],["10.130.6001","Sönmüş kireç CL 70S","ton",0.128,3059.0],["10.130.9991","Su","m³",0.29,55.0],["10.100.1062","Düz işçi","Sa",3.0,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.15,2550.0]]],
[1017,"19.100.2523","0.200 m³/150 Kg kireç-çimento karışımı kaba harç yapılması (sönmüş kireç torbalı)","m³",1966.88,1966.88,"Makine/İşçilik","",[["10.130.1005","","m³",1.0,340.0],["10.130.1209","","ton",0.15,2550.0],["10.130.6001","Sönmüş kireç CL 70S","ton",0.15,3059.0],["10.130.9991","Su","m³",0.305,55.0],["10.100.1062","Düz işçi","Sa",2.75,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0],["10.240.3201","Mermer pirinci (beyaz)","ton",1.45,420.0],["10.130.1209","","ton",0.65,2550.0]]],
[1018,"19.100.2525","Mozayik harcı yapılması","m³",3205.5,3205.5,"Makine/İşçilik","",[["10.240.3201","Mermer pirinci (beyaz)","ton",1.45,420.0],["10.130.1209","","ton",0.65,2550.0],["10.130.9991","Su","m³",0.3,55.0],["10.100.1062","Düz işçi","Sa",3.5,205.0],["10.100.1062","Düz işçi","Sa",1.0,205.0]]],
[1019,"19.100.3002","yollar için)","Sa",7979.44,7979.44,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.25,355.0],["10.100.1057","Operatör yardımcısı","Sa",2.5,300.0],["10.100.1059","Yağcı","Sa",1.25,230.0]]],
[1020,"19.100.3003","Kayar Kalıplı Beton Finişerinin (Beton Yol İçin) Bir Saatlik Ücreti","Sa",37189.8,37189.8,"Makine/İşçilik","",[["10.100.1055","Operatör makinist","Sa",1.35,355.0],["10.100.1057","Operatör yardımcısı","Sa",2.7,300.0],["10.100.1059","Yağcı","Sa",1.35,230.0]]]
];
const POZLAR_BASLANGIC = _PD.map(d=>({
  id:d[0],no:d[1],tanim:d[2],birim:d[3],
  birimFiyat:d[5],          // Mayıs (varsayılan gösterim)
  ocakFiyat:d[4],           // Ocak (analiz)
  mayisFiyat:d[5],          // Mayıs (resmi liste)
  idare:"Çevre ve Şehircilik",kategori:d[6],kaynak:"2026",
  aciklama:d[7],
  rayicler:d[8].map(r=>({no:r[0],tanim:r[1],birim:r[2],miktar:r[3],birimFiyat:r[4]}))
}));

// Mutable global poz store — PDF import sonrası güncellenir
const _store = { pozlar: POZLAR_BASLANGIC };
const getPozlar = () => _store.pozlar;

const KAT_RENK = {
  "Beton İşleri":"#3b82f6","Demir İşleri":"#ef4444","Kaba İnşaat":"#f59e0b",
  "İnce İnşaat":"#10b981","Tesisat":"#8b5cf6","Elektrik":"#f97316","Zemin İşleri":"#6b7280",
  "Nakliye":"#0ea5e9","Diğer":"#94a3b8",
};

const BOYUTLAR = {
  "m³":[{key:"en",label:"En (m)"},{key:"boy",label:"Boy (m)"},{key:"yukseklik",label:"Yükseklik (m)"}],
  "m²":[{key:"en",label:"En (m)"},{key:"boy",label:"Boy (m)"}],
  "m": [{key:"uzunluk",label:"Uzunluk (m)"}],
  "kg":[{key:"uzunluk",label:"Uzunluk (m)"},{key:"en",label:"Çap/En (mm)"}],
  "ton":[{key:"uzunluk",label:"Uzunluk (m)"},{key:"en",label:"Genişlik (m)"}],
};

const fmt = (n) => Number(n||0).toLocaleString("tr-TR",{minimumFractionDigits:2});
const uid = () => Math.random().toString(36).slice(2,8);

const boyutHesapla = (birim, b) => {
  const e=parseFloat(b.en||0), y=parseFloat(b.boy||0), h=parseFloat(b.yukseklik||0), u=parseFloat(b.uzunluk||0);
  if(birim==="m³") return e*y*h;
  if(birim==="m²") return e*y;
  if(birim==="m")  return u;
  return u;
};

// ── Veri yapısı ──
// projeler: [{id, ad, renk, bolumler:[{id,ad,pozlar:[{pozId,satirlar:[...]}]}]}]

const RENKLER = ["#6366f1","#f59e0b","#10b981","#ef4444","#8b5cf6","#f97316","#3b82f6"];

const basProje = () => ({ id:uid(), ad:"Yeni Proje", renk:RENKLER[0], bolumler:[] });
const basBolum = () => ({ id:uid(), ad:"Yeni Bölüm", pozlar:[] });
const basPozKayit = (pozId) => ({ pozId, satirlar:[] });
const basSatir = (minha=false) => ({
  id:uid(), aciklama:"", girisTipi:"boyut",
  boyutlar:{en:"",boy:"",yukseklik:"",uzunluk:""}, miktar:"", carpan:"1", minha
});

// ── Hesap yardımcıları ──
const satirSonuc = (satir, birim) => {
  const m = satir.girisTipi==="boyut" ? boyutHesapla(birim, satir.boyutlar) : parseFloat(satir.miktar||0);
  return m * parseFloat(satir.carpan||1);
};
const pozNetMetraj = (pozKayit, birim) =>
  pozKayit.satirlar.reduce((a,s) => a + (s.minha ? -1 : 1) * satirSonuc(s,birim), 0);
const pozBrut = (pozKayit, birim) =>
  pozKayit.satirlar.filter(s=>!s.minha).reduce((a,s)=>a+satirSonuc(s,birim),0);
const pozMinha = (pozKayit, birim) =>
  pozKayit.satirlar.filter(s=>s.minha).reduce((a,s)=>a+satirSonuc(s,birim),0);

const bolumToplam = (bolum) =>
  bolum.pozlar.reduce((a,pk)=>{
    const poz=getPozlar().find(p=>p.id===pk.pozId);
    return a + (poz ? pozNetMetraj(pk,poz.birim)*poz.birimFiyat : 0);
  },0);

// ── Storage yardımcıları ──────────────────────────────────────────
const storageGet = async (key) => {
  try { const r = await window.storage.get(key); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
};
const storageSet = async (key, val) => {
  try { await window.storage.set(key, JSON.stringify(val)); } catch {}
};

const projeToplam = (proje) => proje.bolumler.reduce((a,b)=>a+bolumToplam(b),0);

// ── Storage başlangıç verisi (render öncesi sync okunur) ──────────
// Async storage'dan veri gelene kadar başlangıç state'i
let _initProjeler = null;
let _initPozlar = null;
let _initPozKaynagi = null;

// ── Ana uygulama ──────────────────────────────────────────────────
export default function App() {
  const pozlarRef = useRef(POZLAR_BASLANGIC);
  const [pozlar, _setPozlar] = useState(POZLAR_BASLANGIC);
  const setPozlar = (v) => {
    const yeni = typeof v === "function" ? v(pozlarRef.current) : v;
    pozlarRef.current = yeni;
    _store.pozlar = yeni;
    _setPozlar(yeni);
    storageSet("mx:pozlar", yeni);
  };
  const pozKaynagiRef = useRef({ yil: 2026, ay: "Mayıs", tip: "Resmi Liste" });
  const [pozKaynagi, _setPozKaynagi] = useState({ yil: 2026, ay: "Mayıs", tip: "Resmi Liste" });
  const setPozKaynagi = (v) => { pozKaynagiRef.current = v; _setPozKaynagi(v); storageSet("mx:pozKaynagi", v); };

  const DEMO_PROJE = [{ id:"p1", ad:"Zemin Kat İnşaatı", renk:"#6366f1",
    bolumler:[
      { id:"b1", ad:"Betonarme İşleri", pozlar:[
        { pozId:1, satirlar:[{id:"s1",aciklama:"Temel plak",girisTipi:"boyut",boyutlar:{en:"10",boy:"12",yukseklik:"0.3",uzunluk:""},miktar:"",carpan:"1",minha:false}] },
      ]},
      { id:"b2", ad:"Sıva & Boya", pozlar:[
        { pozId:8, satirlar:[
          {id:"s2",aciklama:"Salon duvarı",girisTipi:"boyut",boyutlar:{en:"5",boy:"3",yukseklik:"",uzunluk:""},miktar:"",carpan:"4",minha:false},
          {id:"s3",aciklama:"Kapı boşlukları",girisTipi:"boyut",boyutlar:{en:"0.9",boy:"2.1",yukseklik:"",uzunluk:""},miktar:"",carpan:"3",minha:true}
        ]},
      ]},
    ]
  }];

  const [projeler, _setProjeler] = useState(DEMO_PROJE);
  const setProjeler = (v) => {
    const yeni = typeof v === "function" ? v(projeler) : v;
    _setProjeler(yeni);
    storageSet("mx:projeler", yeni);
  };
  const [aktifProjeId, setAktifProjeId] = useState("p1");
  const [yukleniyorStorage, setYukleniyorStorage] = useState(false);
  const storageYuklendi = useRef(false);

  // Storage yükle - Promise.then ile render sonrası güncelle
  if (!storageYuklendi.current) {
    storageYuklendi.current = true;
    setTimeout(() => {
      Promise.all([
        storageGet("mx:projeler"),
        storageGet("mx:pozlar"),
        storageGet("mx:pozKaynagi"),
      ]).then(([sp, sk, spk]) => {
        if (sp && sp.length > 0) { _setProjeler(sp); setAktifProjeId(sp[0]?.id || "p1"); }
        if (sk && sk.length > 0) { pozlarRef.current = sk; _store.pozlar = sk; _setPozlar(sk); }
        if (spk) { pozKaynagiRef.current = spk; _setPozKaynagi(spk); }
      }).catch(() => {});
    }, 0);
  }

  const [anaEkran, setAnaEkran] = useState("projeler");
  const [aramaMetni, setAramaMetni] = useState("");
  const [seciliIdare, setSeciliIdare] = useState("Tümü");
  const [seciliKat, setSeciliKat] = useState("Tümü");
  const [pozEklemeHedef, setPozEklemeHedef] = useState(null);
  const [projeModalAc, setProjeModalAc] = useState(false);
  const [bolumModalAc, setBolumModalAc] = useState(null);
  const [yeniProjeAd, setYeniProjeAd] = useState("");
  const [yeniProjeRenk, setYeniProjeRenk] = useState(RENKLER[0]);
  const [yeniBolumAd, setYeniBolumAd] = useState("");

  // 2. Silme onay modal
  const [onayModal, setOnayModal] = useState(null); // {mesaj, onAy, tip}
  const silOnay = (mesaj, onAy) => setOnayModal({ mesaj, onAy });

  const aktifProje = projeler.find(p=>p.id===aktifProjeId) || projeler[0];
  const genelToplam = projeler.reduce((a,p)=>a+projeToplam(p),0);
  const aktifToplam = aktifProje ? projeToplam(aktifProje) : 0;

  // ── Proje işlemleri ──
  const projeEkle = () => {
    if(!yeniProjeAd.trim()) return;
    const yeni = { ...basProje(), ad:yeniProjeAd.trim(), renk:yeniProjeRenk };
    setProjeler(p=>[...p, yeni]);
    setAktifProjeId(yeni.id);
    setYeniProjeAd(""); setProjeModalAc(false);
  };
  const projeAdGuncelle = (id, ad) => setProjeler(p=>p.map(x=>x.id===id?{...x,ad}:x));
  const projeSil = (id) => {
    silOnay(`"${projeler.find(p=>p.id===id)?.ad}" projesini ve tüm verilerini silmek istediğinize emin misiniz?`, () => {
      setProjeler(p=>p.filter(x=>x.id!==id));
      setAktifProjeId(prev=>prev===id ? (projeler.find(p=>p.id!==id)?.id||null) : prev);
    });
  };

  const bolumEkle = (projeId, ad) => {
    if(!ad || !ad.trim()) return;
    const yeni = { ...basBolum(), ad:ad.trim() };
    setProjeler(p=>p.map(x=>x.id===projeId ? {...x, bolumler:[...x.bolumler,yeni]} : x));
  };
  const bolumAdGuncelle = (projeId, bolumId, ad) =>
    setProjeler(p=>p.map(x=>x.id===projeId?{...x,bolumler:x.bolumler.map(b=>b.id===bolumId?{...b,ad}:b)}:x));
  const bolumSil = (projeId, bolumId) => {
    const proje = projeler.find(p=>p.id===projeId);
    const bolum = proje?.bolumler.find(b=>b.id===bolumId);
    silOnay(`"${bolum?.ad}" bölümünü ve içindeki tüm pozları silmek istediğinize emin misiniz?`, () =>
      setProjeler(p=>p.map(x=>x.id===projeId?{...x,bolumler:x.bolumler.filter(b=>b.id!==bolumId)}:x))
    );
  };

  // ── Poz işlemleri ──
  const pozEkle = (projeId, bolumId, pozId) => {
    setProjeler(p=>p.map(proje=>{
      if(proje.id!==projeId) return proje;
      return {...proje, bolumler:proje.bolumler.map(bolum=>{
        if(bolum.id!==bolumId) return bolum;
        if(bolum.pozlar.find(pk=>pk.pozId===pozId)) {
          // zaten var, sadece satır ekle
          return {...bolum, pozlar:bolum.pozlar.map(pk=>pk.pozId===pozId?{...pk,satirlar:[...pk.satirlar,basSatir()]}:pk)};
        }
        return {...bolum, pozlar:[...bolum.pozlar,{...basPozKayit(pozId),satirlar:[basSatir()]}]};
      })};
    }));
    setAnaEkran("metraj");
    setPozEklemeHedef(null);
  };

  // ── Satır işlemleri ──
  const satirGuncelle = (projeId, bolumId, pozId, satirId, alan, deger) =>
    setProjeler(p=>p.map(proje=>proje.id!==projeId?proje:{...proje,bolumler:proje.bolumler.map(bolum=>bolum.id!==bolumId?bolum:{...bolum,pozlar:bolum.pozlar.map(pk=>pk.pozId!==pozId?pk:{...pk,satirlar:pk.satirlar.map(s=>s.id!==satirId?s:{...s,[alan]:deger})})})}));

  const boyutGuncelle = (projeId, bolumId, pozId, satirId, alan, deger) =>
    setProjeler(p=>p.map(proje=>proje.id!==projeId?proje:{...proje,bolumler:proje.bolumler.map(bolum=>bolum.id!==bolumId?bolum:{...bolum,pozlar:bolum.pozlar.map(pk=>pk.pozId!==pozId?pk:{...pk,satirlar:pk.satirlar.map(s=>s.id!==satirId?s:{...s,boyutlar:{...s.boyutlar,[alan]:deger}})})})}));

  const satirEkle = (projeId, bolumId, pozId, minha=false) =>
    setProjeler(p=>p.map(proje=>proje.id!==projeId?proje:{...proje,bolumler:proje.bolumler.map(bolum=>bolum.id!==bolumId?bolum:{...bolum,pozlar:bolum.pozlar.map(pk=>pk.pozId!==pozId?pk:{...pk,satirlar:[...pk.satirlar,basSatir(minha)]})})}));

  const satirSil = (projeId, bolumId, pozId, satirId) =>
    setProjeler(p=>p.map(proje=>proje.id!==projeId?proje:{...proje,bolumler:proje.bolumler.map(bolum=>bolum.id!==bolumId?bolum:{...bolum,pozlar:bolum.pozlar.map(pk=>pk.pozId!==pozId?pk:{...pk,satirlar:pk.satirlar.filter(s=>s.id!==satirId)})})}));

  // ── Excel export ──
  const excelExport = () => {
    const wb = XLSX.utils.book_new();
    const tarih = new Date().toLocaleDateString("tr-TR");

    // Başlık stili yardımcısı
    const baslik = (metin) => [[metin],[`Oluşturma Tarihi: ${tarih}`],[]];

    // ── 1. Metraj Cetveli ──
    if(secilenRaporlar.metrajCetveli) {
      projeler.forEach(proje => {
        const rows = [...baslik(`METRAJ CETVELİ — ${proje.ad.toUpperCase()}`)];
        proje.bolumler.forEach(bolum => {
          rows.push([`BÖLÜM: ${bolum.ad}`,`Bölüm Toplamı: ₺${fmt(bolumToplam(bolum))}`,"","","","",""]);
          bolum.pozlar.forEach(pk => {
            const poz = getPozlar().find(p=>p.id===pk.pozId); if(!poz) return;
            rows.push([`  POZ: ${poz.no} — ${poz.tanim}`,"","","","","",""]);
            rows.push(["  Açıklama","Giriş Tipi","Boyut 1","Boyut 2","Boyut 3","Çarpan",`Sonuç (${poz.birim})`]);
            pk.satirlar.forEach(s=>{
              const al=BOYUTLAR[poz.birim]||[{key:"uzunluk"}];
              const son=satirSonuc(s,poz.birim);
              rows.push([
                (s.minha?"[MİNHA] ":"") + (s.aciklama||""),
                s.girisTipi==="boyut"?"Boyut":"Direkt",
                s.girisTipi==="boyut"?(parseFloat(s.boyutlar[al[0]?.key]||0)||""):(parseFloat(s.miktar||0)||""),
                s.girisTipi==="boyut"?(parseFloat(s.boyutlar[al[1]?.key]||0)||""):"",
                s.girisTipi==="boyut"?(parseFloat(s.boyutlar[al[2]?.key]||0)||""):"",
                parseFloat(s.carpan||1), s.minha?-son:son,
              ]);
            });
            const brut=pozBrut(pk,poz.birim), minha=pozMinha(pk,poz.birim), net=pozNetMetraj(pk,poz.birim);
            rows.push(["","","","","","Brüt Metraj:", brut]);
            if(minha>0) rows.push(["","","","","","Minha (−):", -minha]);
            rows.push(["","","","","","Net Metraj:", net]);
            rows.push(["","","","","","Birim Fiyat (₺):", poz.birimFiyat]);
            rows.push(["","","","","","TUTAR (₺):", net*poz.birimFiyat]);
            rows.push([]);
          });
          rows.push([]);
        });
        const ws=XLSX.utils.aoa_to_sheet(rows);
        ws["!cols"]=[{wch:38},{wch:10},{wch:10},{wch:10},{wch:10},{wch:16},{wch:14}];
        XLSX.utils.book_append_sheet(wb, ws, `Metraj-${proje.ad.slice(0,20)}`);
      });
    }

    // ── 2. Yaklaşık Maliyet Hesap Cetveli ──
    if(secilenRaporlar.maliyetCetveli) {
      const rows = [...baslik("YAKLAŞIK MALİYET HESAP CETVELİ")];
      rows.push(["Proje","Bölüm","Poz No","İş Kaleminin Adı ve Kısa Tanımı","Birimi","Miktarı","Birim Fiyatı (₺)","Tutarı (₺)"]);
      projeler.forEach(proje=>{
        proje.bolumler.forEach(bolum=>{
          bolum.pozlar.forEach(pk=>{
            const poz=getPozlar().find(p=>p.id===pk.pozId); if(!poz) return;
            const net=pozNetMetraj(pk,poz.birim);
            rows.push([proje.ad, bolum.ad, poz.no, poz.tanim, poz.birim, net, poz.birimFiyat, net*poz.birimFiyat]);
          });
        });
      });
      rows.push([]);
      rows.push(["","","","","","","KDV Hariç Toplam (₺)", genelToplam]);
      rows.push(["","","","","","","KDV (%20) (₺)", genelToplam*0.20]);
      rows.push(["","","","","","","KDV Dahil Genel Toplam (₺)", genelToplam*1.20]);
      const ws=XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"]=[{wch:20},{wch:18},{wch:12},{wch:40},{wch:8},{wch:12},{wch:16},{wch:16}];
      XLSX.utils.book_append_sheet(wb, ws, "YM Hesap Cetveli");
    }

    // ── 3. Yaklaşık Maliyet İcmal Tablosu ──
    if(secilenRaporlar.maliyetIcmal) {
      const rows = [...baslik("YAKLAŞIK MALİYET İCMAL TABLOSU")];
      rows.push(["Sıra","Proje / Bölüm","Tutar (₺)","Oran (%)"]);
      let sira=1;
      projeler.forEach(proje=>{
        const pt=projeToplam(proje);
        rows.push([`${sira++}`, proje.ad.toUpperCase(), pt, genelToplam>0?((pt/genelToplam)*100).toFixed(2)+"%":"-"]);
        proje.bolumler.forEach(bolum=>{
          const bt=bolumToplam(bolum);
          rows.push(["", `  └ ${bolum.ad}`, bt, pt>0?((bt/pt)*100).toFixed(2)+"%":"-"]);
        });
      });
      rows.push([]);
      rows.push(["","GENEL TOPLAM (KDV Hariç)", genelToplam, "100.00%"]);
      rows.push(["","KDV (%20)", genelToplam*0.20, ""]);
      rows.push(["","KDV DAHİL TOPLAM", genelToplam*1.20, ""]);
      const ws=XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"]=[{wch:6},{wch:36},{wch:18},{wch:12}];
      XLSX.utils.book_append_sheet(wb, ws, "YM İcmal Tablosu");
    }

    // ── 4. Pursantaj Tablosu ──
    if(secilenRaporlar.pursantaj) {
      const rows = [...baslik("PURSANTAJ TABLOSU")];
      rows.push(["Sıra","Poz No","İş Kaleminin Adı","Birimi","Net Miktar","Birim Fiyat (₺)","Tutar (₺)","Pursantaj (%)"]);
      let sira=1;
      projeler.forEach(proje=>{
        proje.bolumler.forEach(bolum=>{
          bolum.pozlar.forEach(pk=>{
            const poz=getPozlar().find(p=>p.id===pk.pozId); if(!poz) return;
            const net=pozNetMetraj(pk,poz.birim);
            const tutar=net*poz.birimFiyat;
            const pursantaj=genelToplam>0?((tutar/genelToplam)*100).toFixed(4):0;
            rows.push([sira++, poz.no, poz.tanim, poz.birim, net, poz.birimFiyat, tutar, pursantaj+"%"]);
          });
        });
      });
      rows.push([]);
      rows.push(["","","","","","","TOPLAM (₺)", genelToplam]);
      rows.push(["","","","","","","TOPLAM (%)", "100.0000%"]);
      const ws=XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"]=[{wch:6},{wch:12},{wch:40},{wch:8},{wch:12},{wch:16},{wch:16},{wch:14}];
      XLSX.utils.book_append_sheet(wb, ws, "Pursantaj Tablosu");
    }

    // ── 5. Kategori Dağılım Analizi ──
    if(secilenRaporlar.kategoriAnaliz) {
      const katMap = {};
      projeler.forEach(proje=>proje.bolumler.forEach(bolum=>bolum.pozlar.forEach(pk=>{
        const poz=getPozlar().find(p=>p.id===pk.pozId); if(!poz) return;
        const t=pozNetMetraj(pk,poz.birim)*poz.birimFiyat;
        katMap[poz.kategori]=(katMap[poz.kategori]||0)+t;
      })));
      const rows = [...baslik("KATEGORİ DAĞILIM ANALİZİ")];
      rows.push(["Sıra","Kategori","Tutar (₺)","Oran (%)","Bar"]);
      const sorted=Object.entries(katMap).sort(([,a],[,b])=>b-a);
      sorted.forEach(([kat,tutar],i)=>{
        const oran=genelToplam>0?((tutar/genelToplam)*100):0;
        const bar="█".repeat(Math.round(oran/2));
        rows.push([i+1, kat, tutar, oran.toFixed(2)+"%", bar]);
      });
      rows.push([]);
      rows.push(["","TOPLAM", genelToplam, "100.00%", ""]);
      const ws=XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"]=[{wch:6},{wch:22},{wch:18},{wch:12},{wch:30}];
      XLSX.utils.book_append_sheet(wb, ws, "Kategori Analizi");
    }

    // ── 6. KDV Hesap Tablosu ──
    if(secilenRaporlar.kdvTablosu) {
      const rows = [...baslik("KDV HESAP TABLOSU")];
      rows.push(["Sıra","Proje","KDV Hariç (₺)","KDV Matrahı (%20) (₺)","KDV Dahil Toplam (₺)"]);
      projeler.forEach((proje,i)=>{
        const pt=projeToplam(proje);
        rows.push([i+1, proje.ad, pt, pt*0.20, pt*1.20]);
      });
      rows.push([]);
      rows.push(["","GENEL TOPLAM", genelToplam, genelToplam*0.20, genelToplam*1.20]);
      const ws=XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"]=[{wch:6},{wch:28},{wch:18},{wch:22},{wch:22}];
      XLSX.utils.book_append_sheet(wb, ws, "KDV Hesap Tablosu");
    }

    // ── 7. Proje Karşılaştırma ──
    if(secilenRaporlar.projeKarsilastirma) {
      const rows = [...baslik("PROJE KARŞILAŞTIRMA TABLOSU")];
      rows.push(["Sıra","Proje Adı","Bölüm Sayısı","Poz Sayısı","Tutar (₺)","Oran (%)","KDV Dahil (₺)"]);
      projeler.forEach((proje,i)=>{
        const pt=projeToplam(proje);
        const toplamPoz=proje.bolumler.reduce((a,b)=>a+b.pozlar.length,0);
        rows.push([i+1, proje.ad, proje.bolumler.length, toplamPoz, pt, genelToplam>0?((pt/genelToplam)*100).toFixed(2)+"%":"-", pt*1.20]);
      });
      rows.push([]);
      const toplamPoz=projeler.reduce((a,p)=>a+p.bolumler.reduce((b,bo)=>b+bo.pozlar.length,0),0);
      rows.push(["","TOPLAM", projeler.length+" proje", toplamPoz+" poz", genelToplam, "100.00%", genelToplam*1.20]);
      const ws=XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"]=[{wch:6},{wch:28},{wch:14},{wch:12},{wch:18},{wch:12},{wch:18}];
      XLSX.utils.book_append_sheet(wb, ws, "Proje Karşılaştırma");
    }

    if(wb.SheetNames.length===0) { alert("Lütfen en az bir rapor seçin."); return; }
    const projeAdlari = projeler.map(p=>p.ad.replace(/[^a-zA-Z0-9ğüşıöçĞÜŞİÖÇ]/g,"_").slice(0,15)).join("-");
    XLSX.writeFile(wb,`MetrajX_${projeAdlari}_${tarih.replace(/\./g,"-")}.xlsx`);
    setExcelModalAc(false);
  };

  // ── Oska uyumlu Excel export ──
  const oskaExport = () => {
    const wb = XLSX.utils.book_new();
    const tarih = new Date().toLocaleDateString("tr-TR");

    projeler.forEach(proje => {
      const rows = [];
      // Oska'nın beklediği başlık satırı — ilk 3 sütun: Poz No, Tanımı, Birimi
      rows.push(["Poz No", "Tanımı", "Birimi", "Birim Fiyatı", "Miktarı", "Tutarı"]);

      proje.bolumler.forEach(bolum => {
        // Bölüm başlığı satırı (Oska bunu açıklama satırı olarak okur)
        rows.push([bolum.ad, "", "", "", "", ""]);

        bolum.pozlar.forEach(pk => {
          const poz = getPozlar().find(p => p.id === pk.pozId);
          if (!poz) return;
          const net = pozNetMetraj(pk, poz.birim);
          const tutar = net * poz.birimFiyat;

          // Oska formatı: virgüllü ondalık (Türkçe), nokta yok
          const fmtOska = (n) => Number(n||0).toLocaleString("tr-TR", {
            minimumFractionDigits: 2, maximumFractionDigits: 2,
            useGrouping: false  // basamak grubu yok — Oska bunu ister
          });

          rows.push([
            poz.no,           // Poz No — Metin olmalı
            poz.tanim,        // Tanımı
            poz.birim,        // Birimi
            fmtOska(poz.birimFiyat),  // Birim Fiyatı — virgüllü
            fmtOska(net),             // Miktarı — virgüllü
            fmtOska(tutar),           // Tutarı
          ]);
        });
      });

      // Toplam satırı
      const toplamTutar = proje.bolumler.reduce((a,b)=>a+bolumToplam(b),0);
      rows.push(["", "TOPLAM", "", "", "", Number(toplamTutar).toLocaleString("tr-TR",{minimumFractionDigits:2,useGrouping:false})]);

      const ws = XLSX.utils.aoa_to_sheet(rows);

      // Poz No sütunu Metin formatında olmalı (Oska şartı)
      const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
      for (let r = 1; r <= range.e.r; r++) {
        const cell = ws[XLSX.utils.encode_cell({r, c: 0})];
        if (cell) cell.t = "s"; // string olarak işaretle
      }

      ws["!cols"] = [{wch:14},{wch:50},{wch:8},{wch:14},{wch:12},{wch:14}];
      XLSX.utils.book_append_sheet(wb, ws, proje.ad.slice(0,28));
    });

    XLSX.writeFile(wb, `MetrajX_Oska_${tarih.replace(/\./g,"-")}.xlsx`);
  };

  const [excelModalAc, setExcelModalAc] = useState(false);
  const [secilenRaporlar, setSecilenRaporlar] = useState({
    metrajCetveli: true,
    maliyetCetveli: true,
    maliyetIcmal: true,
    pursantaj: true,
    kategoriAnaliz: true,
    kdvTablosu: true,
    projeKarsilastirma: true,
  });

  const RAPOR_LISTESI = [
    { key: "metrajCetveli",      icon: "📐", ad: "Metraj Cetveli",              aciklama: "Poz bazlı satır satır brüt / minha / net metraj" },
    { key: "maliyetCetveli",     icon: "📋", ad: "Yaklaşık Maliyet Hesap Cetveli", aciklama: "Poz no, tarif, birim, miktar, birim fiyat, tutar" },
    { key: "maliyetIcmal",       icon: "🗂", ad: "Yaklaşık Maliyet İcmal Tablosu", aciklama: "Proje ve bölüm bazlı toplam tutarlar" },
    { key: "pursantaj",          icon: "📊", ad: "Pursantaj Tablosu",            aciklama: "Her iş kaleminin genel toplama oranı (%)" },
    { key: "kategoriAnaliz",     icon: "🥧", ad: "Kategori Dağılım Analizi",     aciklama: "Kategori bazlı tutar ve yüzde dağılımı" },
    { key: "kdvTablosu",         icon: "🧾", ad: "KDV Hesap Tablosu",            aciklama: "KDV hariç, KDV tutarı, KDV dahil kırılım (proje bazlı)" },
    { key: "projeKarsilastirma", icon: "⚖️", ad: "Proje Karşılaştırma",          aciklama: "Projeler arası tutar ve pursantaj karşılaştırması" },
  ];

  const tumunuSec = () => setSecilenRaporlar(Object.fromEntries(RAPOR_LISTESI.map(r=>[r.key,true])));
  const tumunuKaldir = () => setSecilenRaporlar(Object.fromEntries(RAPOR_LISTESI.map(r=>[r.key,false])));
  const hepsiSecili = RAPOR_LISTESI.every(r=>secilenRaporlar[r.key]);

  const filtreliPozlar = useMemo(()=>pozlar.filter(p=>{
    const q=aramaMetni.toLowerCase();
    return (p.tanim.toLowerCase().includes(q)||p.no.toLowerCase().includes(q))
      &&(seciliIdare==="Tümü"||p.idare===seciliIdare)
      &&(seciliKat==="Tümü"||p.kategori===seciliKat);
  }),[pozlar, aramaMetni,seciliIdare,seciliKat]);

  if (yukleniyorStorage) return (
    <div style={{...S.root,alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}>
      <div style={{fontSize:36}}>⚙</div>
      <div style={{fontSize:16,fontWeight:700,color:"#60a5fa"}}>MetrajX yükleniyor...</div>
      <div style={{fontSize:12,color:"#475569"}}>Veriler geri yükleniyor</div>
    </div>
  );

  const [sidebarDar, setSidebarDar] = useState(false);

  return (
    <div style={S.root}>
      {onayModal && (
        <div style={S.overlay} onClick={()=>setOnayModal(null)}>
          <div style={{...S.modal,width:380,textAlign:"center"}} onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:36,marginBottom:12}}>⚠️</div>
            <div style={{fontSize:14,fontWeight:700,color:"#f1f5f9",marginBottom:8}}>Emin misiniz?</div>
            <div style={{fontSize:13,color:"#94a3b8",marginBottom:20,lineHeight:1.6}}>{onayModal.mesaj}</div>
            <div style={{display:"flex",gap:8}}>
              <button style={{...S.ikinciBtn,flex:1}} onClick={()=>setOnayModal(null)}>İptal</button>
              <button style={{flex:1,padding:"9px 0",backgroundColor:"#7f1d1d",border:"1px solid #ef4444",borderRadius:8,color:"#fca5a5",fontSize:13,fontWeight:700,cursor:"pointer"}}
                onClick={()=>{onayModal.onAy();setOnayModal(null);}}>Evet, Sil</button>
            </div>
          </div>
        </div>
      )}

      {/* ── SOL SIDEBAR ── */}
      <aside style={{...S.sidebar, width: sidebarDar ? 52 : 220, transition:"width 0.2s"}}>
        {/* Logo + daralt butonu */}
        <div style={{display:"flex",alignItems:"center",justifyContent: sidebarDar?"center":"space-between",padding: sidebarDar?"12px 0":"12px 14px 12px",borderBottom:"1px solid #1e293b",marginBottom:8}}>
          {!sidebarDar && <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:20}}>⚙</span>
            <div style={S.logoTitle}>MetrajX</div>
          </div>}
          {sidebarDar && <span style={{fontSize:20}}>⚙</span>}
          <button onClick={()=>setSidebarDar(d=>!d)}
            style={{background:"transparent",border:"1px solid #1e293b",borderRadius:6,color:"#475569",cursor:"pointer",fontSize:12,padding:"3px 6px",flexShrink:0}}>
            {sidebarDar?"▶":"◀"}
          </button>
        </div>

        {/* Nav */}
        <div style={{padding:"0 6px"}}>
          {[["projeler","🏗","Projeler"],["kutuphane","📚","Poz Kütüphanesi"],["pdfimport","📤","Poz Güncelle"]].map(([key,icon,label])=>(
            <button key={key} title={label}
              style={{...S.navItem, justifyContent: sidebarDar?"center":"flex-start", padding: sidebarDar?"10px 0":"9px 10px",
                ...(anaEkran===key?S.navItemAktif:{})}
              } onClick={()=>setAnaEkran(key)}>
              <span style={{fontSize:16}}>{icon}</span>
              {!sidebarDar && <span style={{flex:1,fontSize:13}}>{label}</span>}
              {!sidebarDar && key==="pdfimport" && <span style={{fontSize:9,backgroundColor:"#1e3a5f",color:"#60a5fa",padding:"2px 4px",borderRadius:3,fontWeight:700}}>YENİ</span>}
            </button>
          ))}
        </div>

        <div style={{flex:1}}/>

        {/* Poz kaynağı - sadece geniş modda */}
        {!sidebarDar && (
          <div style={{padding:"6px 10px",backgroundColor:"#050a14",margin:"0 6px 6px",borderRadius:6,border:"1px solid #1e293b"}}>
            <div style={{fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase",marginBottom:2}}>Poz Kaynağı</div>
            <div style={{fontSize:11,color:"#60a5fa",fontWeight:600}}>{pozKaynagi.tip} {pozKaynagi.yil}{pozKaynagi.ay?` / ${pozKaynagi.ay}`:""}</div>
            <div style={{fontSize:10,color:"#334155"}}>{pozlar.length} poz</div>
          </div>
        )}
        <button style={{...S.excelBtn,margin: sidebarDar?"6px 4px":"6px 6px", fontSize: sidebarDar?16:12, padding: sidebarDar?"8px 0":"8px 0"}}
          title="Excel İndir" onClick={()=>setExcelModalAc(true)}>
          {sidebarDar ? "📥" : "📥 Excel İndir"}
        </button>
      </aside>

      {/* ── ANA İÇERİK ── */}
      <main style={S.main}>
        {anaEkran==="projeler" && (
          <ProjelerEkrani
            projeler={projeler} aktifProjeId={aktifProjeId} setAktifProjeId={setAktifProjeId}
            setProjeler={setProjeler} projeEkle={projeEkle} projeSil={projeSil}
            projeAdGuncelle={projeAdGuncelle} bolumEkle={bolumEkle} bolumSil={bolumSil}
            bolumAdGuncelle={bolumAdGuncelle} satirGuncelle={satirGuncelle}
            boyutGuncelle={boyutGuncelle} satirEkle={satirEkle} satirSil={satirSil}
            setPozEklemeHedef={setPozEklemeHedef} setAnaEkran={setAnaEkran}
            yeniProjeAd={yeniProjeAd} setYeniProjeAd={setYeniProjeAd}
            yeniProjeRenk={yeniProjeRenk} setYeniProjeRenk={setYeniProjeRenk}
            genelToplam={genelToplam} excelExport={excelExport}
          />
        )}
        {anaEkran==="kutuphane" && (
          <KutuphaneEkrani pozlar={filtreliPozlar}
            aramaMetni={aramaMetni} setAramaMetni={setAramaMetni}
            seciliIdare={seciliIdare} setSeciliIdare={setSeciliIdare}
            seciliKat={seciliKat} setSeciliKat={setSeciliKat}
            pozEklemeHedef={pozEklemeHedef} setPozEklemeHedef={setPozEklemeHedef}
            pozEkle={pozEkle} aktifProje={aktifProje} setAnaEkran={setAnaEkran}
          />
        )}
        {anaEkran==="pdfimport" && (
          <PdfImportEkrani
            setPozlar={setPozlar} setPozKaynagi={setPozKaynagi}
            mevcutPozSayisi={pozlar.length} pozKaynagi={pozKaynagi}
            setAnaEkran={setAnaEkran}
          />
        )}

        {/* EXCEL MODAL */}
        {excelModalAc && (
          <div style={S.overlay} onClick={()=>setExcelModalAc(false)}>
            <div style={{...S.modal,width:520}} onClick={e=>e.stopPropagation()}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                <div style={S.modalBaslik}>📥 Excel Çıktı Seçenekleri</div>
                <button style={{background:"transparent",border:"none",color:"#475569",cursor:"pointer",fontSize:18}} onClick={()=>setExcelModalAc(false)}>✕</button>
              </div>
              <div style={{fontSize:12,color:"#475569",marginBottom:16}}>İndirilecek rapor sayfalarını seçin</div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",backgroundColor:"#0a0f1e",borderRadius:8,padding:"8px 12px",marginBottom:12,border:"1px solid #1e293b"}}>
                <span style={{fontSize:13,color:"#94a3b8",fontWeight:600}}>{RAPOR_LISTESI.filter(r=>secilenRaporlar[r.key]).length} / {RAPOR_LISTESI.length} rapor seçili</span>
                <button style={{...S.ikinciBtn,padding:"4px 12px",fontSize:12}} onClick={hepsiSecili?tumunuKaldir:tumunuSec}>{hepsiSecili?"Tümünü Kaldır":"Tümünü Seç"}</button>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:20}}>
                {RAPOR_LISTESI.map(rapor=>{
                  const secili=secilenRaporlar[rapor.key];
                  return (
                    <div key={rapor.key} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 12px",borderRadius:8,
                      border:`1px solid ${secili?"#3b82f6":"#1e293b"}`,backgroundColor:secili?"#0d1f3c":"#111827",cursor:"pointer"}}
                      onClick={()=>setSecilenRaporlar(p=>({...p,[rapor.key]:!p[rapor.key]}))}>
                      <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${secili?"#3b82f6":"#334155"}`,backgroundColor:secili?"#3b82f6":"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                        {secili && <span style={{color:"#fff",fontSize:12,fontWeight:900}}>✓</span>}
                      </div>
                      <span style={{fontSize:18}}>{rapor.icon}</span>
                      <div style={{flex:1}}>
                        <div style={{fontSize:13,fontWeight:700,color:secili?"#e2e8f0":"#64748b"}}>{rapor.ad}</div>
                        <div style={{fontSize:11,color:"#475569",marginTop:2}}>{rapor.aciklama}</div>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Oska ayırıcı */}
              <div style={{borderTop:"1px solid #1e293b",margin:"16px 0 12px",paddingTop:12}}>
                <div style={{fontSize:11,color:"#475569",fontWeight:700,textTransform:"uppercase",marginBottom:8}}>
                  🔗 Oska HakedişPlus'a Aktar
                </div>
                <div style={{fontSize:11,color:"#334155",marginBottom:10,lineHeight:1.6}}>
                  Oska'nın kabul ettiği formatta çıktı alır: Poz No · Tanımı · Birimi · Birim Fiyatı · Miktarı sütunları. Her proje ayrı sayfa olarak gelir. Oska'da <b style={{color:"#94a3b8"}}>Excel'den Getir</b> ile içe aktar.
                </div>
                <button style={{width:"100%",padding:"10px 0",backgroundColor:"#1e1a3c",border:"1px solid #6366f1",borderRadius:8,color:"#a5b4fc",fontSize:13,fontWeight:700,cursor:"pointer"}}
                  onClick={()=>{oskaExport();setExcelModalAc(false);}}>
                  🟣 Oska Uyumlu Excel İndir
                </button>
              </div>

              <div style={{display:"flex",gap:8}}>
                <button style={{...S.birincilBtn,flex:1,backgroundColor:"#052e16",borderColor:"#16a34a",color:"#4ade80",fontSize:14}} onClick={excelExport}>
                  📥 {RAPOR_LISTESI.filter(r=>secilenRaporlar[r.key]).length} Sayfa İndir
                </button>
                <button style={{...S.ikinciBtn}} onClick={()=>setExcelModalAc(false)}>İptal</button>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

// ── Projeler Ana Ekranı ──────────────────────────────────────────
function ProjelerEkrani({projeler,aktifProjeId,setAktifProjeId,setProjeler,projeEkle,projeSil,
  projeAdGuncelle,bolumEkle,bolumSil,bolumAdGuncelle,satirGuncelle,boyutGuncelle,
  satirEkle,satirSil,setPozEklemeHedef,setAnaEkran,
  yeniProjeAd,setYeniProjeAd,yeniProjeRenk,setYeniProjeRenk,
  genelToplam, excelExport}) {

  const [projeModalAc, setProjeModalAc] = useState(false);
  const [bolumModalAc, setBolumModalAc] = useState(null);
  const [yeniBolumAd, setYeniBolumAd] = useState("");
  const [aktifSekme, setAktifSekme] = useState("metraj"); // metraj | ozet

  const aktifProje = projeler.find(p=>p.id===aktifProjeId) || projeler[0];

  return (
    <div style={{display:"flex",height:"100%",overflow:"hidden"}}>
      {/* Sol panel - proje listesi */}
      <div style={{width:240,borderRight:"1px solid #1e293b",display:"flex",flexDirection:"column",backgroundColor:"#0d1424",flexShrink:0}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"14px 14px 10px"}}>
          <div style={{fontSize:13,fontWeight:800,color:"#f1f5f9"}}>Projeler</div>
          <button style={{...S.birincilBtn,padding:"4px 12px",fontSize:12}} onClick={()=>setProjeModalAc(true)}>+ Yeni</button>
        </div>

        <div style={{flex:1,overflowY:"auto",padding:"0 8px"}}>
          {projeler.length===0 && (
            <div style={{padding:20,textAlign:"center",color:"#475569",fontSize:12}}>Henüz proje yok.</div>
          )}
          {projeler.map(proje=>{
            const aktif = aktifProjeId===proje.id;
            return (
              <div key={proje.id}
                style={{borderRadius:10,padding:"11px 12px",marginBottom:6,cursor:"pointer",
                  border:`1px solid ${aktif?proje.renk:"#1e293b"}`,
                  backgroundColor:aktif?"#0d1f3c":"#111827",transition:"all 0.15s"}}
                onClick={()=>setAktifProjeId(proje.id)}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}>
                  <span style={{width:9,height:9,borderRadius:"50%",backgroundColor:proje.renk,flexShrink:0}}/>
                  <span style={{fontSize:13,fontWeight:700,color:aktif?"#f1f5f9":"#94a3b8",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                    {proje.ad}
                  </span>
                  {aktif && projeler.length>1 && (
                    <button style={{background:"transparent",border:"none",color:"#475569",cursor:"pointer",fontSize:12,padding:0,flexShrink:0}}
                      onClick={e=>{e.stopPropagation();projeSil(proje.id);}}>🗑</button>
                  )}
                </div>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:11}}>
                  <span style={{color:"#475569"}}>{proje.bolumler.length} bölüm · {proje.bolumler.reduce((a,b)=>a+b.pozlar.length,0)} poz</span>
                  <span style={{color:aktif?"#10b981":"#475569",fontWeight:600}}>₺{fmt(projeToplam(proje)).split(",")[0]}</span>
                </div>
              </div>
            );
          })}
        </div>

        {/* Genel toplam */}
        <div style={{padding:"10px 12px",borderTop:"1px solid #1e293b"}}>
          <div style={{fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase",marginBottom:3}}>Genel Toplam</div>
          <div style={{fontSize:15,fontWeight:800,color:"#10b981"}}>₺{fmt(genelToplam)}</div>
        </div>
      </div>

      {/* Sağ panel - proje içeriği */}
      <div style={{flex:1,overflow:"auto",display:"flex",flexDirection:"column"}}>
        {!aktifProje ? (
          <div style={{...S.bosEkran,flex:1}}>
            <div style={{fontSize:48,marginBottom:12}}>🏗</div>
            <div style={{fontSize:16,fontWeight:600,marginBottom:6}}>Proje seçin veya oluşturun</div>
            <button style={{...S.birincilBtn,marginTop:8}} onClick={()=>setProjeModalAc(true)}>+ Yeni Proje</button>
          </div>
        ) : (
          <>
            {/* Proje üst bar */}
            <div style={{borderBottom:"1px solid #1e293b",backgroundColor:"#0d1424",padding:"0 20px",display:"flex",alignItems:"center",gap:0,flexShrink:0}}>
              <span style={{width:10,height:10,borderRadius:"50%",backgroundColor:aktifProje.renk,marginRight:10}}/>
              <span style={{fontSize:15,fontWeight:800,color:"#f1f5f9",marginRight:20}}>{aktifProje.ad}</span>
              {/* Proje içi sekmeler */}
              {[["metraj","📐 Metraj"],["ozet","📊 Özet"]].map(([key,label])=>(
                <button key={key} style={{padding:"12px 14px",background:"transparent",border:"none",
                  borderBottom: aktifSekme===key?"2px solid #3b82f6":"2px solid transparent",
                  color: aktifSekme===key?"#60a5fa":"#64748b",cursor:"pointer",fontSize:13,
                  fontWeight:aktifSekme===key?700:500,marginBottom:-1}}
                  onClick={()=>setAktifSekme(key)}>{label}</button>
              ))}
              <div style={{flex:1}}/>
              <span style={{fontSize:13,color:"#10b981",fontWeight:700}}>₺{fmt(projeToplam(aktifProje))}</span>
            </div>

            {/* İçerik */}
            <div style={{flex:1,overflow:"auto"}}>
              {aktifSekme==="metraj" && (
                <MetrajEkrani proje={aktifProje}
                  satirGuncelle={satirGuncelle} boyutGuncelle={boyutGuncelle}
                  satirEkle={satirEkle} satirSil={satirSil}
                  bolumEkle={()=>setBolumModalAc(aktifProje.id)}
                  bolumSil={bolumSil} bolumAdGuncelle={bolumAdGuncelle}
                  projeAdGuncelle={projeAdGuncelle} projeSil={projeSil}
                  setPozEklemeHedef={setPozEklemeHedef} setAnaEkran={setAnaEkran}
                />
              )}
              {aktifSekme==="ozet" && (
                <OzetEkrani projeler={[aktifProje]} genelToplam={projeToplam(aktifProje)} excelExport={excelExport} />
              )}
            </div>
          </>
        )}
      </div>

      {/* Yeni Proje Modal */}
      {projeModalAc && (
        <div style={S.overlay} onClick={()=>setProjeModalAc(false)}>
          <div style={S.modal} onClick={e=>e.stopPropagation()}>
            <div style={S.modalBaslik}>Yeni Proje</div>
            <label style={S.formEtiket}>Proje Adı</label>
            <input style={{...S.formInput,marginBottom:16}} placeholder="Ör: A Blok Kaba İnşaat"
              value={yeniProjeAd} onChange={e=>setYeniProjeAd(e.target.value)}
              onKeyDown={e=>e.key==="Enter"&&(projeEkle(),setProjeModalAc(false))} autoFocus />
            <label style={{...S.formEtiket,marginBottom:8}}>Renk</label>
            <div style={{display:"flex",gap:8,marginBottom:20}}>
              {["#6366f1","#f59e0b","#10b981","#ef4444","#8b5cf6","#f97316","#3b82f6"].map(r=>(
                <button key={r} onClick={()=>setYeniProjeRenk(r)}
                  style={{width:28,height:28,borderRadius:"50%",backgroundColor:r,
                    border:yeniProjeRenk===r?"3px solid #fff":"2px solid transparent",cursor:"pointer"}}/>
              ))}
            </div>
            <div style={{display:"flex",gap:8}}>
              <button style={{...S.birincilBtn,flex:1}} onClick={()=>{projeEkle();setProjeModalAc(false);}}>Oluştur</button>
              <button style={{...S.ikinciBtn,flex:1}} onClick={()=>setProjeModalAc(false)}>İptal</button>
            </div>
          </div>
        </div>
      )}

      {/* Yeni Bölüm Modal */}
      {bolumModalAc && (
        <div style={S.overlay} onClick={()=>setBolumModalAc(null)}>
          <div style={S.modal} onClick={e=>e.stopPropagation()}>
            <div style={S.modalBaslik}>Yeni Bölüm</div>
            <label style={S.formEtiket}>Bölüm Adı</label>
            <input style={S.formInput} placeholder="Ör: Betonarme İşleri"
              value={yeniBolumAd} onChange={e=>setYeniBolumAd(e.target.value)}
              onKeyDown={e=>e.key==="Enter"&&(bolumEkle(bolumModalAc,yeniBolumAd),setYeniBolumAd(""),setBolumModalAc(null))} autoFocus />
            <div style={{display:"flex",gap:8,marginTop:16}}>
              <button style={{...S.birincilBtn,flex:1}} onClick={()=>{bolumEkle(bolumModalAc,yeniBolumAd);setYeniBolumAd("");setBolumModalAc(null);}}>Oluştur</button>
              <button style={{...S.ikinciBtn,flex:1}} onClick={()=>setBolumModalAc(null)}>İptal</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


// ── Metraj Ekranı ─────────────────────────────────────────────────
function MetrajEkrani({proje,satirGuncelle,boyutGuncelle,satirEkle,satirSil,bolumEkle,bolumSil,bolumAdGuncelle,projeAdGuncelle,projeSil,setPozEklemeHedef,setAnaEkran}) {
  const [duzenleAd, setDuzenleAd] = useState(null); // bolumId veya "proje"
  const [adDeger, setAdDeger] = useState("");

  const adKaydet = (tip, id) => {
    if(tip==="proje") projeAdGuncelle(proje.id, adDeger);
    else bolumAdGuncelle(proje.id, id, adDeger);
    setDuzenleAd(null);
  };

  return (
    <div style={S.ekran}>
      {/* Proje başlığı */}
      <div style={{...S.ekranBaslik, borderLeft:`4px solid ${proje.renk}`, paddingLeft:16, marginBottom:24}}>
        {duzenleAd==="proje" ? (
          <input style={{...S.formInput,fontSize:18,fontWeight:800,flex:1}} value={adDeger}
            onChange={e=>setAdDeger(e.target.value)}
            onBlur={()=>adKaydet("proje")} onKeyDown={e=>e.key==="Enter"&&adKaydet("proje")} autoFocus />
        ) : (
          <h1 style={{...S.h1,cursor:"pointer"}} onClick={()=>{setDuzenleAd("proje");setAdDeger(proje.ad);}}>
            {proje.ad} <span style={{fontSize:14,color:"#475569"}}>✎</span>
          </h1>
        )}
        <div style={{display:"flex",gap:8}}>
          <div style={S.toplamKutu}>
            <div style={S.miniEtiket}>Proje Toplamı</div>
            <div style={{fontSize:18,fontWeight:800,color:"#10b981"}}>₺{fmt(projeToplam(proje))}</div>
          </div>
        </div>
      </div>

      {proje.bolumler.length===0 && (
        <div style={S.bosEkran}>
          <div style={{fontSize:48,marginBottom:12}}>📂</div>
          <div style={{fontSize:16,fontWeight:600,marginBottom:6}}>Bu projede henüz bölüm yok</div>
          <div style={{fontSize:13,marginBottom:20}}>Başlamak için bir bölüm oluşturun</div>
          <button style={S.birincilBtn} onClick={bolumEkle}>+ Bölüm Ekle</button>
        </div>
      )}

      {proje.bolumler.map(bolum=>(
        <div key={bolum.id} style={S.bolumBlok}>
          {/* Bölüm başlığı */}
          <div style={S.bolumBaslik}>
            <div style={{flex:1}}>
              {duzenleAd===bolum.id ? (
                <input style={{...S.formInput,fontSize:14,fontWeight:700}} value={adDeger}
                  onChange={e=>setAdDeger(e.target.value)}
                  onBlur={()=>adKaydet("bolum",bolum.id)}
                  onKeyDown={e=>e.key==="Enter"&&adKaydet("bolum",bolum.id)} autoFocus />
              ) : (
                <span style={{fontSize:15,fontWeight:700,color:"#f1f5f9",cursor:"pointer"}}
                  onClick={()=>{setDuzenleAd(bolum.id);setAdDeger(bolum.ad);}}>
                  {bolum.ad} <span style={{fontSize:12,color:"#475569"}}>✎</span>
                </span>
              )}
              <span style={{fontSize:12,color:"#475569",marginLeft:12}}>{bolum.pozlar.length} poz</span>
            </div>
            <span style={{fontSize:14,color:"#10b981",fontWeight:700}}>₺{fmt(bolumToplam(bolum))}</span>
            <button style={S.tehlikeBtn} title="Bölümü Sil" onClick={()=>bolumSil(proje.id,bolum.id)}>🗑</button>
          </div>

          {/* Pozlar */}
          {bolum.pozlar.map(pk=>{
            const poz=getPozlar().find(p=>p.id===pk.pozId); if(!poz) return null;
            const alanlar=BOYUTLAR[poz.birim]||[{key:"uzunluk",label:"Miktar"}];
            const net=pozNetMetraj(pk,poz.birim);
            const brut=pozBrut(pk,poz.birim);
            const minha=pozMinha(pk,poz.birim);
            return (
              <div key={pk.pozId} style={S.pozBlok}>
                <div style={S.pozBlokBaslik}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontFamily:"monospace",fontSize:12,color:"#6366f1",fontWeight:700}}>{poz.no}</span>
                    <span style={{...S.chip,backgroundColor:(KAT_RENK[poz.kategori]||"#475569")+"18",color:KAT_RENK[poz.kategori]||"#475569"}}>{poz.kategori}</span>
                  </div>
                  <div style={{fontSize:13,color:"#cbd5e1",marginTop:4,marginBottom:8}}>{poz.tanim}</div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <span style={{fontSize:12,color:"#475569"}}>Birim: {poz.birim} &nbsp;|&nbsp; ₺{fmt(poz.birimFiyat)}/{poz.birim}</span>
                    <div style={{display:"flex",gap:12,alignItems:"center"}}>
                      {minha>0 && <span style={{fontSize:12,color:"#ef4444"}}>Brüt: {fmt(brut)} — Minha: {fmt(minha)}</span>}
                      <span style={{fontSize:13,color:"#10b981",fontWeight:700}}>Net: {fmt(net)} {poz.birim}</span>
                      <span style={{fontSize:13,color:"#60a5fa",fontWeight:700}}>= ₺{fmt(net*poz.birimFiyat)}</span>
                    </div>
                  </div>
                </div>

                {/* Tablo başlık */}
                <div style={S.cetvelBaslik}>
                  <span style={{flex:2}}>Açıklama / Mahal</span>
                  <span style={{flex:1,textAlign:"center"}}>Giriş</span>
                  {alanlar.map(a=><span key={a.key} style={{flex:1,textAlign:"center"}}>{a.label}</span>)}
                  <span style={{flex:0.8,textAlign:"center"}}>Çarpan</span>
                  <span style={{flex:1,textAlign:"center"}}>Sonuç ({poz.birim})</span>
                  <span style={{width:56}}/>
                </div>

                {/* Satırlar */}
                {pk.satirlar.map(satir=>{
                  const son=satirSonuc(satir,poz.birim);
                  const isMinha=satir.minha;
                  return (
                    <div key={satir.id} style={{...S.cetvelSatir,...(isMinha?S.cetvelSatirMinha:{})}}>
                      <span style={{flex:2,display:"flex",alignItems:"center",gap:6}}>
                        {isMinha&&<span style={S.minhaEtiket}>MİNHA</span>}
                        <input style={{...S.satirInput,...(isMinha?{borderColor:"#7f1d1d",backgroundColor:"#1a0a0a"}:{})}}
                          placeholder={isMinha?"Kapı/pencere boşluğu...":"Açıklama / mahal..."}
                          value={satir.aciklama}
                          onChange={e=>satirGuncelle(proje.id,bolum.id,pk.pozId,satir.id,"aciklama",e.target.value)}/>
                      </span>
                      <span style={{flex:1,display:"flex",justifyContent:"center"}}>
                        <div style={S.toggle}>
                          <button style={{...S.toggleBtn,...(satir.girisTipi==="boyut"?(isMinha?S.toggleBtnMinha:S.toggleBtnAktif):{})}}
                            onClick={()=>satirGuncelle(proje.id,bolum.id,pk.pozId,satir.id,"girisTipi","boyut")}>Boyut</button>
                          <button style={{...S.toggleBtn,...(satir.girisTipi==="direkt"?(isMinha?S.toggleBtnMinha:S.toggleBtnAktif):{})}}
                            onClick={()=>satirGuncelle(proje.id,bolum.id,pk.pozId,satir.id,"girisTipi","direkt")}>Direkt</button>
                        </div>
                      </span>
                      {satir.girisTipi==="boyut"
                        ? alanlar.map(a=>(
                            <span key={a.key} style={{flex:1,display:"flex",justifyContent:"center"}}>
                              <input type="number" style={{...S.sayiInput,...(isMinha?{borderColor:"#7f1d1d"}:{})}}
                                placeholder="0" value={satir.boyutlar[a.key]||""}
                                onChange={e=>boyutGuncelle(proje.id,bolum.id,pk.pozId,satir.id,a.key,e.target.value)}/>
                            </span>))
                        : <span style={{flex:alanlar.length,display:"flex",justifyContent:"center"}}>
                            <input type="number" style={{...S.sayiInput,width:"70%",...(isMinha?{borderColor:"#7f1d1d"}:{})}}
                              placeholder="Miktar" value={satir.miktar||""}
                              onChange={e=>satirGuncelle(proje.id,bolum.id,pk.pozId,satir.id,"miktar",e.target.value)}/>
                          </span>}
                      <span style={{flex:0.8,display:"flex",justifyContent:"center"}}>
                        <input type="number" style={{...S.sayiInput,...(isMinha?{borderColor:"#7f1d1d"}:{})}}
                          placeholder="1" value={satir.carpan||""}
                          onChange={e=>satirGuncelle(proje.id,bolum.id,pk.pozId,satir.id,"carpan",e.target.value)}/>
                      </span>
                      <span style={{flex:1,textAlign:"center",fontWeight:700,fontSize:13,
                        color:isMinha?"#ef4444":(son>0?"#10b981":"#475569")}}>
                        {son>0?(isMinha?`-${fmt(son)}`:fmt(son)):"—"}
                      </span>
                      <span style={{width:56,display:"flex",gap:4,justifyContent:"center"}}>
                        <button title={isMinha?"Normal yap":"Minha yap"}
                          style={{...S.minhaToggleBtn,...(isMinha?S.minhaToggleBtnAktif:{})}}
                          onClick={()=>satirGuncelle(proje.id,bolum.id,pk.pozId,satir.id,"minha",!isMinha)}>
                          {isMinha?"−":"M"}
                        </button>
                        <button style={S.silBtn} onClick={()=>satirSil(proje.id,bolum.id,pk.pozId,satir.id)}>✕</button>
                      </span>
                    </div>
                  );
                })}

                {/* 5. Poz ara toplam satırı */}
                {pk.satirlar.length > 0 && (
                  <div style={{display:"flex",gap:8,padding:"6px 14px",backgroundColor:"#050a14",borderTop:"1px solid #1e3a5f",alignItems:"center"}}>
                    <span style={{flex:2+1+alanlar.length*1+0.8,fontSize:11,color:"#475569",textAlign:"right",paddingRight:8}}>
                      {pozMinha(pk,poz.birim)>0
                        ? <span>Brüt: <b style={{color:"#94a3b8"}}>{fmt(pozBrut(pk,poz.birim))}</b> − Minha: <b style={{color:"#ef4444"}}>{fmt(pozMinha(pk,poz.birim))}</b> =</span>
                        : <span>Ara Toplam:</span>}
                    </span>
                    <span style={{flex:1,textAlign:"center",fontWeight:800,color:"#60a5fa",fontSize:13}}>
                      {fmt(pozNetMetraj(pk,poz.birim))} {poz.birim}
                    </span>
                    <span style={{width:56}}/>
                  </div>
                )}
                <div style={{display:"flex",borderTop:"1px dashed #1e293b"}}>
                  <button style={{...S.satirEkleBtn,flex:1,borderRight:"1px dashed #1e293b"}}
                    onClick={()=>satirEkle(proje.id,bolum.id,pk.pozId)}>+ Satır Ekle</button>
                  <button style={{...S.satirEkleBtn,flex:1,color:"#ef4444"}}
                    onClick={()=>satirEkle(proje.id,bolum.id,pk.pozId,true)}>− Minha Ekle</button>
                </div>
              </div>
            );
          })}

          {/* Bölüme poz ekle */}
          <button style={S.pozEkleBtn} onClick={()=>{setPozEklemeHedef({projeId:proje.id,bolumId:bolum.id});setAnaEkran("kutuphane");}}>
            + Poz Ekle
          </button>
        </div>
      ))}

      {proje.bolumler.length>0 && (
        <button style={{...S.birincilBtn,marginTop:8}} onClick={bolumEkle}>+ Yeni Bölüm Ekle</button>
      )}
    </div>
  );
}

// ── Kütüphane Ekranı ─────────────────────────────────────────────
function KutuphaneEkrani({pozlar,aramaMetni,setAramaMetni,seciliIdare,setSeciliIdare,seciliKat,setSeciliKat,pozEklemeHedef,setPozEklemeHedef,pozEkle,aktifProje,setAnaEkran}) {
  const [detayPoz, setDetayPoz] = useState(null);

  return (
    <div style={S.ekran}>
      <div style={S.ekranBaslik}>
        <div>
          <h1 style={S.h1}>Poz Kütüphanesi</h1>
          <p style={S.aciklama}>{pozlar.length} poz yüklü</p>
        </div>
        {pozEklemeHedef && (
          <button style={S.ikinciBtn} onClick={()=>{setPozEklemeHedef(null);setAnaEkran("metraj");}}>← Metraj'a Dön</button>
        )}
      </div>

      {pozEklemeHedef && (
        <div style={{backgroundColor:"#0d1f3c",border:"1px solid #3b82f6",borderRadius:10,padding:"10px 16px",
          marginBottom:14,display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:18}}>🎯</span>
          <div>
            <div style={{fontSize:11,color:"#60a5fa",fontWeight:700,textTransform:"uppercase"}}>Seçilen poz şuraya eklenecek:</div>
            <div style={{fontSize:14,fontWeight:700,color:"#e2e8f0"}}>
              {aktifProje?.ad} → {aktifProje?.bolumler.find(b=>b.id===pozEklemeHedef.bolumId)?.ad}
            </div>
          </div>
        </div>
      )}
      {!pozEklemeHedef && (
        <div style={{backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:10,padding:"10px 16px",marginBottom:14}}>
          <div style={{fontSize:12,color:"#475569"}}>💡 Metraj cetvelinde bir bölümün "＋ Poz Ekle" butonuna basarak poz ekleyin — ya da aşağıdan seçip doğrudan ekleyin.</div>
        </div>
      )}

      <div style={S.filtreSatiri}>
        <div style={S.aramaKutu}>
          <span style={{color:"#475569"}}>🔍</span>
          <input style={S.aramaInput} placeholder="Poz no veya tarif..." value={aramaMetni} onChange={e=>setAramaMetni(e.target.value)}/>
          {aramaMetni&&<button style={S.temizleBtn} onClick={()=>setAramaMetni("")}>✕</button>}
        </div>
        <select style={S.select} value={seciliIdare} onChange={e=>setSeciliIdare(e.target.value)}>
          {IDARELER.map(i=><option key={i}>{i}</option>)}
        </select>
        <select style={S.select} value={seciliKat} onChange={e=>setSeciliKat(e.target.value)}>
          {KATEGORILER.map(k=><option key={k}>{k}</option>)}
        </select>
        {(aramaMetni||seciliIdare!=="Tümü"||seciliKat!=="Tümü") && (
          <button style={{...S.ikinciBtn,padding:"7px 12px",fontSize:11}}
            onClick={()=>{setAramaMetni("");setSeciliIdare("Tümü");setSeciliKat("Tümü");}}>
            ✕ Filtreleri Sıfırla
          </button>
        )}
      </div>

      <div style={{fontSize:12,color:"#475569",marginBottom:12}}>{pozlar.length} poz</div>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {pozlar.map(poz=>(
          <div key={poz.id} style={{...S.pozKart,cursor:"pointer"}} onClick={()=>setDetayPoz(poz)}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
              <span style={{fontFamily:"monospace",fontSize:12,color:"#6366f1",fontWeight:700}}>{poz.no}</span>
              <span style={{...S.chip,backgroundColor:(KAT_RENK[poz.kategori]||"#475569")+"18",color:KAT_RENK[poz.kategori]||"#475569"}}>{poz.kategori}</span>
            </div>
            <div style={{fontSize:13,color:"#cbd5e1",lineHeight:1.5,marginBottom:10}}>{poz.tanim}</div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                <span style={{fontSize:11,color:"#475569",backgroundColor:"#1e293b",padding:"2px 8px",borderRadius:4}}>{poz.idare}</span>
                <span style={{fontSize:12,color:"#475569"}}>/ {poz.birim}</span>
                {poz.rayicler?.length > 0 && <span style={{fontSize:10,color:"#334155",backgroundColor:"#1e293b",padding:"1px 6px",borderRadius:4}}>🔍 {poz.rayicler.length} rayiç</span>}
              </div>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <div style={{textAlign:"right"}}>
                  {poz.mayisFiyat && poz.mayisFiyat !== poz.ocakFiyat && (
                    <div style={{fontSize:10,color:"#475569",textDecoration:"line-through"}}>₺{fmt(poz.ocakFiyat)}</div>
                  )}
                  <div style={{fontSize:14,fontWeight:700,color:"#10b981"}}>
                    ₺{fmt(poz.mayisFiyat || poz.birimFiyat)}
                    {poz.mayisFiyat && poz.mayisFiyat !== poz.ocakFiyat && (
                      <span style={{fontSize:9,marginLeft:4,color:"#60a5fa",fontWeight:600}}>MAY</span>
                    )}
                  </div>
                </div>
                {pozEklemeHedef && (
                  <button style={{...S.birincilBtn,padding:"5px 12px",fontSize:12}}
                    onClick={e=>{e.stopPropagation();pozEkle(pozEklemeHedef.projeId,pozEklemeHedef.bolumId,poz.id);}}>
                    + Ekle
                  </button>
                )}
              </div>
            </div>
          </div>
        ))}
      </div>

      {/* Detay Modal */}
      {detayPoz && (
        <div style={S.overlay} onClick={()=>setDetayPoz(null)}>
          <div style={{...S.modal,width:580,maxHeight:"85vh",overflow:"auto"}} onClick={e=>e.stopPropagation()}>
            {/* Başlık */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
              <div style={{flex:1,paddingRight:12}}>
                <div style={{...S.chip,backgroundColor:(KAT_RENK[detayPoz.kategori]||"#475569")+"22",color:KAT_RENK[detayPoz.kategori]||"#475569",display:"inline-block",marginBottom:8}}>{detayPoz.kategori}</div>
                <div style={{fontFamily:"monospace",fontSize:13,color:"#6366f1",fontWeight:700,marginBottom:6}}>{detayPoz.no}</div>
                <div style={{fontSize:15,fontWeight:700,color:"#f1f5f9",lineHeight:1.5}}>{detayPoz.tanim}</div>
              </div>
              <button style={{background:"transparent",border:"none",color:"#475569",cursor:"pointer",fontSize:20,flexShrink:0}} onClick={()=>setDetayPoz(null)}>✕</button>
            </div>

            {/* Birim & Fiyat */}
            <div style={{display:"flex",gap:10,marginBottom:16}}>
              <div style={{flex:1,backgroundColor:"#0a0f1e",borderRadius:8,padding:"10px 12px"}}>
                <div style={{fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Birim</div>
                <div style={{fontSize:13,color:"#e2e8f0"}}>{detayPoz.birim}</div>
              </div>
              <div style={{flex:1.5,backgroundColor:"#052e16",border:"1px solid #16a34a",borderRadius:8,padding:"10px 12px"}}>
                <div style={{fontSize:9,color:"#4ade80",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Mayıs 2026 Fiyatı</div>
                <div style={{fontSize:16,color:"#4ade80",fontWeight:800}}>₺{fmt(detayPoz.mayisFiyat || detayPoz.birimFiyat)}</div>
              </div>
              {detayPoz.ocakFiyat && detayPoz.ocakFiyat !== detayPoz.mayisFiyat && (
                <div style={{flex:1.5,backgroundColor:"#0d1f3c",border:"1px solid #1e3a5f",borderRadius:8,padding:"10px 12px"}}>
                  <div style={{fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Ocak 2026 (Analiz)</div>
                  <div style={{fontSize:13,color:"#94a3b8",fontWeight:700}}>₺{fmt(detayPoz.ocakFiyat)}</div>
                  <div style={{fontSize:9,color: detayPoz.mayisFiyat > detayPoz.ocakFiyat ? "#ef4444" : "#10b981",marginTop:2}}>
                    {detayPoz.mayisFiyat > detayPoz.ocakFiyat ? "▲" : "▼"} %{Math.abs(((detayPoz.mayisFiyat - detayPoz.ocakFiyat)/detayPoz.ocakFiyat)*100).toFixed(1)} fark
                  </div>
                </div>
              )}
              {(!detayPoz.ocakFiyat || detayPoz.ocakFiyat === detayPoz.mayisFiyat) && (
                <div style={{flex:1.5,backgroundColor:"#0a0f1e",borderRadius:8,padding:"10px 12px"}}>
                  <div style={{fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>Ocak 2026 (Analiz)</div>
                  <div style={{fontSize:13,color:"#475569"}}>—</div>
                </div>
              )}
              <div style={{flex:1,backgroundColor:"#0a0f1e",borderRadius:8,padding:"10px 12px"}}>
                <div style={{fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase",marginBottom:4}}>İdare</div>
                <div style={{fontSize:11,color:"#e2e8f0"}}>{detayPoz.idare}</div>
              </div>
            </div>

            {/* Açıklama */}
            {detayPoz.aciklama && (
              <div style={{backgroundColor:"#0d1f3c",border:"1px solid #1e3a5f",borderRadius:8,padding:"10px 14px",marginBottom:14}}>
                <div style={{fontSize:10,color:"#475569",fontWeight:700,textTransform:"uppercase",marginBottom:6}}>📋 Açıklama</div>
                <div style={{fontSize:12,color:"#94a3b8",lineHeight:1.7}}>{detayPoz.aciklama}</div>
              </div>
            )}

            {/* Rayiçler */}
            {detayPoz.rayicler?.length > 0 && (
              <div style={{marginBottom:16}}>
                <div style={{fontSize:10,color:"#475569",fontWeight:700,textTransform:"uppercase",marginBottom:8}}>🔩 Rayiç Analizi ({detayPoz.rayicler.length} kalem)</div>
                <div style={{backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:10,overflow:"hidden"}}>
                  <div style={{display:"flex",gap:8,padding:"7px 12px",backgroundColor:"#0d1424",fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase"}}>
                    <span style={{flex:0.8}}>Poz No</span>
                    <span style={{flex:3}}>Tanımı</span>
                    <span style={{flex:0.6,textAlign:"center"}}>Birim</span>
                    <span style={{flex:0.7,textAlign:"right"}}>Miktar</span>
                    <span style={{flex:1,textAlign:"right"}}>Birim Fiyat</span>
                    <span style={{flex:1,textAlign:"right"}}>Tutar</span>
                  </div>
                  {detayPoz.rayicler.map((r,i)=>(
                    <div key={i} style={{display:"flex",gap:8,padding:"7px 12px",borderTop:"1px solid #1e293b",alignItems:"center"}}>
                      <span style={{flex:0.8,fontFamily:"monospace",fontSize:10,color:"#6366f1"}}>{r.no}</span>
                      <span style={{flex:3,fontSize:11,color:"#cbd5e1"}}>{r.tanim}</span>
                      <span style={{flex:0.6,textAlign:"center",fontSize:10,color:"#94a3b8"}}>{r.birim}</span>
                      <span style={{flex:0.7,textAlign:"right",fontSize:11,color:"#94a3b8"}}>{r.miktar}</span>
                      <span style={{flex:1,textAlign:"right",fontSize:11,color:"#94a3b8"}}>₺{fmt(r.birimFiyat)}</span>
                      <span style={{flex:1,textAlign:"right",fontSize:11,color:"#10b981",fontWeight:600}}>₺{fmt(r.miktar*r.birimFiyat)}</span>
                    </div>
                  ))}
                  {/* Rayiç toplamı */}
                  <div style={{display:"flex",gap:8,padding:"8px 12px",borderTop:"2px solid #1e3a5f",backgroundColor:"#0d1424"}}>
                    <span style={{flex:0.8+3+0.6+0.7+1,color:"#64748b",fontSize:11}}>Malzeme + İşçilik</span>
                    <span style={{flex:1,textAlign:"right",color:"#60a5fa",fontWeight:700,fontSize:12}}>
                      ₺{fmt(detayPoz.rayicler.reduce((a,r)=>a+r.miktar*r.birimFiyat,0))}
                    </span>
                  </div>
                  <div style={{display:"flex",gap:8,padding:"8px 12px",borderTop:"1px solid #1e293b",backgroundColor:"#0a0f1e"}}>
                    <span style={{flex:0.8+3+0.6+0.7+1,color:"#64748b",fontSize:11}}>%25 Yüklenici Kârı + Birim Fiyat</span>
                    <span style={{flex:1,textAlign:"right",color:"#10b981",fontWeight:800,fontSize:14}}>₺{fmt(detayPoz.birimFiyat)}</span>
                  </div>
                </div>
              </div>
            )}

            {/* Ekle butonu */}
            {pozEklemeHedef && (
              <button style={{...S.birincilBtn,width:"100%",padding:"12px 0",fontSize:14,marginTop:4}}
                onClick={()=>{pozEkle(pozEklemeHedef.projeId,pozEklemeHedef.bolumId,detayPoz.id);setDetayPoz(null);}}>
                + Seçilen Bölüme Ekle
              </button>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Özet Ekranı ──────────────────────────────────────────────────
function OzetEkrani({projeler,genelToplam,excelExport}) {
  return (
    <div style={S.ekran}>
      <div style={S.ekranBaslik}>
        <div><h1 style={S.h1}>Genel Özet</h1><p style={S.aciklama}>{projeler.length} proje</p></div>
        <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
          <button style={S.excelBtnBuyuk} onClick={()=>setExcelModalAc(true)}>📥 Excel İndir</button>
          <div style={S.toplamKutu}>
            <div style={S.miniEtiket}>Genel Toplam</div>
            <div style={{fontSize:22,fontWeight:800,color:"#10b981",marginTop:4}}>₺{fmt(genelToplam)}</div>
          </div>
        </div>
      </div>

      <div style={{display:"flex",gap:12,marginBottom:24}}>
        {[["KDV Hariç",genelToplam,"#60a5fa"],["KDV (%20)",genelToplam*0.20,"#f59e0b"],["KDV Dahil",genelToplam*1.20,"#10b981"]].map(([e,d,r])=>(
          <div key={e} style={{flex:1,backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:12,padding:"14px 18px"}}>
            <div style={{fontSize:11,color:"#475569",textTransform:"uppercase",fontWeight:600,marginBottom:6}}>{e}</div>
            <div style={{fontSize:20,fontWeight:800,color:r}}>₺{fmt(d)}</div>
          </div>
        ))}
      </div>

      {projeler.map(proje=>(
        <div key={proje.id} style={{...S.bolumBlok,marginBottom:20}}>
          <div style={{...S.bolumBaslik,borderLeft:`4px solid ${proje.renk}`}}>
            <span style={{flex:1,fontSize:15,fontWeight:700,color:"#f1f5f9"}}>{proje.ad}</span>
            <span style={{color:"#10b981",fontWeight:700}}>₺{fmt(projeToplam(proje))}</span>
          </div>
          {proje.bolumler.map(bolum=>(
            <div key={bolum.id}>
              <div style={{padding:"8px 16px",backgroundColor:"#0d1424",display:"flex",justifyContent:"space-between",fontSize:12,color:"#64748b",fontWeight:600,borderTop:"1px solid #1e293b"}}>
                <span>📁 {bolum.ad}</span>
                <span>₺{fmt(bolumToplam(bolum))}</span>
              </div>
              {bolum.pozlar.map(pk=>{
                const poz=getPozlar().find(p=>p.id===pk.pozId); if(!poz) return null;
                const net=pozNetMetraj(pk,poz.birim);
                const brut=pozBrut(pk,poz.birim);
                const minha=pozMinha(pk,poz.birim);
                return (
                  <div key={pk.pozId} style={{display:"flex",gap:10,padding:"10px 16px",borderTop:"1px solid #1e293b",alignItems:"center"}}>
                    <span style={{fontFamily:"monospace",fontSize:12,color:"#6366f1",flex:1}}>{poz.no}</span>
                    <span style={{flex:4,fontSize:12,color:"#cbd5e1"}}>{poz.tanim}</span>
                    <span style={{flex:0.7,textAlign:"center",color:"#94a3b8",fontSize:12}}>{poz.birim}</span>
                    {minha>0
                      ? <span style={{flex:1.5,textAlign:"right",fontSize:12}}>
                          <span style={{color:"#94a3b8"}}>{fmt(brut)}</span>
                          <span style={{color:"#ef4444"}}> −{fmt(minha)}</span>
                          <span style={{color:"#60a5fa",fontWeight:600}}> = {fmt(net)}</span>
                        </span>
                      : <span style={{flex:1.5,textAlign:"right",color:"#60a5fa",fontWeight:600,fontSize:12}}>{fmt(net)}</span>}
                    <span style={{flex:1.2,textAlign:"right",color:"#94a3b8",fontSize:12}}>₺{fmt(poz.birimFiyat)}</span>
                    <span style={{flex:1.2,textAlign:"right",color:"#10b981",fontWeight:700,fontSize:13}}>₺{fmt(net*poz.birimFiyat)}</span>
                  </div>
                );
              })}
            </div>
          ))}
        </div>
      ))}
    </div>
  );
}

// ── Stiller ─────────────────────────────────────────────────────
const S = {
  root:{display:"flex",height:"100vh",backgroundColor:"#0a0f1e",fontFamily:"'Segoe UI',system-ui,sans-serif",color:"#e2e8f0",overflow:"hidden"},
  sidebar:{width:240,backgroundColor:"#0d1424",borderRight:"1px solid #1e293b",display:"flex",flexDirection:"column",flexShrink:0,overflow:"hidden"},
  logo:{display:"flex",alignItems:"center",gap:10,padding:"16px 16px 14px",borderBottom:"1px solid #1e293b"},
  logoTitle:{fontSize:15,fontWeight:800,color:"#f1f5f9"},logoSub:{fontSize:10,color:"#475569"},
  projelerBaslik:{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 16px 6px"},
  miniEkleBtn:{width:22,height:22,borderRadius:5,border:"1px solid #1e293b",backgroundColor:"#1e293b",color:"#60a5fa",cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"},
  projeListe:{flex:1,overflowY:"auto",padding:"0 8px"},
  projeTab:{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",borderRadius:8,cursor:"pointer",marginBottom:2,borderLeft:"3px solid transparent"},
  projeTabAktif:{backgroundColor:"#111827"},
  navBolum:{padding:"8px 10px",borderTop:"1px solid #1e293b"},
  navItem:{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",borderRadius:8,border:"none",background:"transparent",color:"#64748b",cursor:"pointer",fontSize:13,fontWeight:500,width:"100%",textAlign:"left"},
  navItemAktif:{backgroundColor:"#1e3a5f",color:"#60a5fa",fontWeight:600},
  sidebarAlt:{display:"flex",justifyContent:"space-between",padding:"10px 16px",borderTop:"1px solid #1e293b"},
  miniEtiket:{fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase",marginBottom:2},
  miniDeger:{fontSize:12,fontWeight:800,color:"#60a5fa"},
  excelBtn:{margin:"8px 10px",padding:"8px 0",backgroundColor:"#052e16",border:"1px solid #16a34a",borderRadius:8,color:"#4ade80",cursor:"pointer",fontSize:12,fontWeight:700},
  excelBtnBuyuk:{padding:"9px 18px",backgroundColor:"#052e16",border:"1px solid #16a34a",borderRadius:8,color:"#4ade80",cursor:"pointer",fontSize:13,fontWeight:700},
  main:{flex:1,overflow:"auto"},
  ekran:{padding:24,maxWidth:1100,margin:"0 auto"},
  ekranBaslik:{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20},
  h1:{fontSize:20,fontWeight:800,color:"#f1f5f9",margin:0},aciklama:{fontSize:12,color:"#475569",margin:"3px 0 0"},
  bolumBlok:{backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:12,marginBottom:16,overflow:"hidden"},
  bolumBaslik:{display:"flex",alignItems:"center",gap:12,padding:"12px 16px",backgroundColor:"#0d1424",borderBottom:"1px solid #1e293b"},
  pozBlok:{borderTop:"1px solid #1e293b"},
  pozBlokBaslik:{padding:"12px 16px",backgroundColor:"#0a0f1e"},
  cetvelBaslik:{display:"flex",gap:8,padding:"7px 14px",backgroundColor:"#050a14",fontSize:10,color:"#475569",fontWeight:700,textTransform:"uppercase"},
  cetvelSatir:{display:"flex",gap:8,padding:"7px 14px",alignItems:"center",borderTop:"1px solid #1e293b"},
  cetvelSatirMinha:{backgroundColor:"#0f0505",borderLeft:"3px solid #ef4444"},
  minhaEtiket:{fontSize:9,fontWeight:800,color:"#ef4444",backgroundColor:"#7f1d1d33",padding:"2px 5px",borderRadius:4,whiteSpace:"nowrap"},
  minhaToggleBtn:{width:24,height:24,borderRadius:5,border:"1px solid #1e293b",backgroundColor:"#1e293b",color:"#64748b",cursor:"pointer",fontSize:13,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"},
  minhaToggleBtnAktif:{backgroundColor:"#7f1d1d",borderColor:"#ef4444",color:"#fca5a5"},
  toggle:{display:"flex",backgroundColor:"#1e293b",borderRadius:6,overflow:"hidden"},
  toggleBtn:{padding:"3px 7px",fontSize:11,border:"none",backgroundColor:"transparent",color:"#64748b",cursor:"pointer"},
  toggleBtnAktif:{backgroundColor:"#3b82f6",color:"#fff",fontWeight:700},
  toggleBtnMinha:{backgroundColor:"#7f1d1d",color:"#fca5a5",fontWeight:700},
  satirInput:{width:"100%",backgroundColor:"#1e293b",border:"1px solid #334155",borderRadius:6,color:"#e2e8f0",fontSize:12,padding:"4px 8px",outline:"none"},
  sayiInput:{width:66,backgroundColor:"#1e293b",border:"1px solid #334155",borderRadius:6,color:"#e2e8f0",fontSize:12,padding:"4px 6px",outline:"none",textAlign:"center"},
  satirEkleBtn:{padding:"8px 0",backgroundColor:"transparent",border:"none",color:"#3b82f6",cursor:"pointer",fontSize:12,fontWeight:600},
  pozEkleBtn:{width:"100%",padding:"10px 0",backgroundColor:"transparent",border:"none",borderTop:"2px dashed #1e293b",color:"#6366f1",cursor:"pointer",fontSize:13,fontWeight:700},
  silBtn:{width:24,height:24,background:"transparent",border:"1px solid #1e293b",borderRadius:5,color:"#475569",cursor:"pointer",fontSize:11,flexShrink:0},
  tehlikeBtn:{background:"transparent",border:"none",cursor:"pointer",fontSize:15,padding:"2px 6px",color:"#475569"},
  pozKart:{backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:10,padding:"12px 14px"},
  chip:{fontSize:10,fontWeight:600,padding:"2px 7px",borderRadius:20},
  bosEkran:{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:80,color:"#475569"},
  toplamKutu:{backgroundColor:"#0d1f3c",border:"1px solid #1e3a5f",borderRadius:10,padding:"10px 16px",textAlign:"right"},
  filtreSatiri:{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"},
  aramaKutu:{flex:1,minWidth:180,display:"flex",alignItems:"center",gap:8,backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:8,padding:"0 12px"},
  aramaInput:{flex:1,background:"transparent",border:"none",outline:"none",color:"#e2e8f0",fontSize:13,padding:"8px 0"},
  temizleBtn:{background:"transparent",border:"none",color:"#475569",cursor:"pointer",fontSize:13},
  select:{backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:8,color:"#94a3b8",fontSize:12,padding:"8px 10px",outline:"none",cursor:"pointer"},
  overlay:{position:"fixed",inset:0,backgroundColor:"rgba(0,0,0,0.65)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100},
  modal:{backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:16,padding:26,width:420,position:"relative"},
  modalBaslik:{fontSize:16,fontWeight:800,color:"#f1f5f9",marginBottom:16},
  formEtiket:{fontSize:11,color:"#475569",fontWeight:700,textTransform:"uppercase",display:"block",marginBottom:6},
  formInput:{width:"100%",backgroundColor:"#0a0f1e",border:"1px solid #1e293b",borderRadius:8,color:"#e2e8f0",fontSize:13,padding:"9px 12px",outline:"none",boxSizing:"border-box"},
  birincilBtn:{padding:"9px 18px",backgroundColor:"#1e3a5f",border:"1px solid #3b82f6",borderRadius:8,color:"#60a5fa",fontSize:13,fontWeight:700,cursor:"pointer"},
  ikinciBtn:{padding:"9px 18px",backgroundColor:"#1e293b",border:"1px solid #334155",borderRadius:8,color:"#94a3b8",fontSize:13,fontWeight:600,cursor:"pointer"},
  detayBtn:{fontSize:11,color:"#94a3b8",backgroundColor:"transparent",border:"1px solid #1e293b",borderRadius:6,padding:"4px 10px",cursor:"pointer"},
  tablo:{backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:12,overflow:"hidden"},
  tabloBaslik:{display:"flex",gap:10,padding:"10px 14px",backgroundColor:"#0d1424",fontSize:10,color:"#475569",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.4px"},
  tabloSatir:{display:"flex",gap:10,padding:"10px 14px",alignItems:"center",borderTop:"1px solid #1e293b"},
  tabloToplam:{display:"flex",gap:10,padding:"12px 14px",alignItems:"center",borderTop:"2px solid #1e3a5f",backgroundColor:"#0d1424"},
};

// ── PDF Import Ekranı ─────────────────────────────────────────────
function PdfImportEkrani({ setPozlar, setPozKaynagi, mevcutPozSayisi, pozKaynagi, setAnaEkran }) {
  const [durum, setDurum] = useState("bos"); // bos | yukleniyor | analiz | onay | hata
  const [pdfBase64, setPdfBase64] = useState(null);
  const [dosyaAdi, setDosyaAdi] = useState("");
  const [analizSonucu, setAnalizSonucu] = useState(null);
  const [hata, setHata] = useState("");
  const [surukle, setSurukle] = useState(false);
  const [yil, setYil] = useState("2026");
  const AYLAR_TR = ["Ocak","Şubat","Mart","Nisan","Mayıs","Haziran","Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"];
  const [ay, setAy] = useState("");
  const [listeTipi, setListeTipi] = useState("İnşaat Birim Fiyat");

  const dosyaOku = (file) => {
    if (!file) { setHata("Dosya seçilemedi."); setDurum("hata"); return; }
    // iPad/Mac'te PDF MIME type farklı gelebilir, uzantıyı da kontrol et
    const isPdf = file.type === "application/pdf" || file.type === "application/octet-stream" || file.name.toLowerCase().endsWith(".pdf");
    if (!isPdf) { setHata("Lütfen PDF dosyası seçin."); setDurum("hata"); return; }
    setDosyaAdi(file.name); setDurum("yukleniyor");
    const reader = new FileReader();
    reader.onload = (e) => {
      const result = e.target.result;
      const base64 = result.includes(",") ? result.split(",")[1] : result;
      setPdfBase64(base64);
      setDurum("hazir");
    };
    reader.onerror = (e) => { setHata("Dosya okunamadı: " + (e.target.error?.message || "bilinmeyen hata")); setDurum("hata"); };
    reader.readAsDataURL(file);
  };

  const pdfAnaliz = async () => {
    if (!pdfBase64) return;
    setDurum("analiz");
    setHata("");
    try {
      const sistem = `Sen Türk inşaat sektörü birim fiyat listesi uzmanısın. 
Sana verilen PDF, Çevre, Şehircilik ve İklim Değişikliği Bakanlığı Yüksek Fen Kurulu tarafından yayınlanan resmi birim fiyat listesidir.
PDF'den tüm pozları çıkar ve SADECE aşağıdaki JSON formatında yanıt ver, başka hiçbir şey yazma:
{
  "yil": "2026",
  "ay": "Ocak veya null",
  "tip": "İnşaat Birim Fiyat veya Mekanik veya Elektrik",
  "pozlar": [
    {
      "no": "poz numarası (örn: 16.001/1)",
      "tanim": "pozun tam tanımı",
      "birim": "m², m³, m, kg, ton, adet vb.",
      "birimFiyat": 1234.56
    }
  ]
}
Kategorileri şu gruplardan birine ata: Beton İşleri, Demir İşleri, Kaba İnşaat, İnce İnşaat, Tesisat, Elektrik, Zemin İşleri, Nakliye, Diğer.
birimFiyat sadece sayı olsun (TL simgesi ve virgül olmadan, nokta ondalık ayracı).
Eğer birim fiyat aylık güncelleme listesindeyse o güncel fiyatı kullan.
Maksimum 200 poz çıkar, en önemli ve sık kullanılanları önceliklendir.`;

      const resp = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 8000,
          system: sistem,
          messages: [{
            role: "user",
            content: [{
              type: "document",
              source: { type: "base64", media_type: "application/pdf", data: pdfBase64 }
            }, {
              type: "text",
              text: `Bu PDF'den birim fiyat pozlarını çıkar. Yıl: ${yil}, ${ay ? "Ay: " + ay + "," : ""} Liste tipi: ${listeTipi}. SADECE JSON döndür.`
            }]
          }]
        })
      });

      if (!resp.ok) throw new Error(`API hatası: ${resp.status}`);
      const data = await resp.json();
      const rawText = data.content.map(c => c.text || "").join("");
      
      // JSON temizle
      const jsonMatch = rawText.match(/\{[\s\S]*\}/);
      if (!jsonMatch) throw new Error("PDF'den poz çıkarılamadı. Farklı bir sayfa deneyin.");
      
      const parsed = JSON.parse(jsonMatch[0]);
      if (!parsed.pozlar || parsed.pozlar.length === 0) throw new Error("Bu PDF'de poz bulunamadı.");
      
      // Kategori ata
      const kategoriMap = {
        "16.":"Beton İşleri","21.":"Demir İşleri","23.":"Kaba İnşaat","27.":"İnce İnşaat",
        "08.":"Tesisat","09.":"Elektrik","12.":"Zemin İşleri","13.":"Nakliye"
      };
      parsed.pozlar = parsed.pozlar.map((p, i) => {
        let kat = "Diğer";
        for (const [prefix, k] of Object.entries(kategoriMap)) {
          if (p.no && p.no.startsWith(prefix)) { kat = k; break; }
        }
        if (listeTipi.includes("Elektrik")) kat = "Elektrik";
        if (listeTipi.includes("Mekanik") || listeTipi.includes("Tesisat")) kat = "Tesisat";
        return {
          id: Date.now() + i,
          no: p.no || "?",
          tanim: p.tanim || "",
          birim: p.birim || "adet",
          birimFiyat: parseFloat(p.birimFiyat) || 0,
          idare: "Çevre ve Şehircilik",
          kategori: kat,
          kaynak: `${parsed.yil || yil}${parsed.ay ? " " + parsed.ay : ""}`,
        };
      });

      setAnalizSonucu(parsed);
      setDurum("onay");
    } catch (err) {
      setHata(err.message || "Beklenmeyen hata.");
      setDurum("hata");
    }
  };

  const pozlariUygula = (mod) => {
    // mod: "yenile" = tamamen değiştir, "ekle" = mevcut + yeni
    const yeniPozlar = analizSonucu.pozlar;
    if (mod === "yenile") {
      setPozlar(yeniPozlar);
    } else {
      // Mevcut pozlarda olmayanları ekle, olanları fiyat güncelle
      setPozlar(prev => {
        const noMap = new Map(prev.map(p => [p.no, p]));
        const guncellenecek = [];
        const eklenecek = [];
        yeniPozlar.forEach(yp => {
          if (noMap.has(yp.no)) guncellenecek.push(yp);
          else eklenecek.push(yp);
        });
        const guncellendi = prev.map(p => {
          const yeni = yeniPozlar.find(yp => yp.no === p.no);
          return yeni ? { ...p, birimFiyat: yeni.birimFiyat, kaynak: yeni.kaynak } : p;
        });
        return [...guncellendi, ...eklenecek];
      });
    }
    setPozKaynagi({
      yil: parseInt(yil),
      ay: ay || null,
      tip: listeTipi,
    });
    setDurum("tamam");
  };

  return (
    <div style={S.ekran}>
      <div style={S.ekranBaslik}>
        <div>
          <h1 style={S.h1}>📤 Poz Güncelleme</h1>
          <p style={S.aciklama}>Bakanlık PDF'inden poz kütüphanesini güncelle</p>
        </div>
        <div style={{backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:10,padding:"10px 16px",textAlign:"right"}}>
          <div style={{fontSize:10,color:"#475569",fontWeight:700,textTransform:"uppercase"}}>Mevcut Kaynak</div>
          <div style={{fontSize:13,color:"#60a5fa",fontWeight:700}}>{pozKaynagi.tip} {pozKaynagi.yil}{pozKaynagi.ay ? " / " + pozKaynagi.ay : ""}</div>
          <div style={{fontSize:11,color:"#475569"}}>{mevcutPozSayisi} poz</div>
        </div>
      </div>

      {/* Adım göstergesi */}
      <div style={{display:"flex",gap:0,marginBottom:28,backgroundColor:"#111827",borderRadius:10,overflow:"hidden",border:"1px solid #1e293b"}}>
        {[["1","PDF Yükle",["bos","hazir","yukleniyor"]],["2","Analiz Et",["analiz"]],["3","Onayla",["onay","tamam"]]].map(([num,lab,aktifler],i)=>{
          const aktif = aktifler.includes(durum) || (num==="3" && durum==="tamam");
          const tamamlandi = (num==="1" && ["analiz","onay","tamam"].includes(durum)) || (num==="2" && ["onay","tamam"].includes(durum));
          return (
            <div key={num} style={{flex:1,padding:"12px 0",textAlign:"center",backgroundColor:tamamlandi?"#052e16":aktif?"#0d1f3c":"transparent",borderRight:i<2?"1px solid #1e293b":"none"}}>
              <div style={{fontSize:18,marginBottom:3}}>{tamamlandi?"✅":aktif?"🔵":"⚪"}</div>
              <div style={{fontSize:11,fontWeight:700,color:tamamlandi?"#4ade80":aktif?"#60a5fa":"#475569"}}>{num}. {lab}</div>
            </div>
          );
        })}
      </div>

      {/* Bilgi kartı */}
      <div style={{backgroundColor:"#0d1f3c",border:"1px solid #1e3a5f",borderRadius:10,padding:"14px 18px",marginBottom:20}}>
        <div style={{fontSize:12,color:"#60a5fa",fontWeight:700,marginBottom:6}}>📌 Nasıl Kullanılır?</div>
        <div style={{fontSize:12,color:"#94a3b8",lineHeight:1.7}}>
          1. <a href="https://yfk.csb.gov.tr/birim-fiyatlar-100468" target="_blank" style={{color:"#60a5fa"}}>yfk.csb.gov.tr</a> adresinden birim fiyat PDF'ini indir<br/>
          2. Aylık güncelleme için <a href="https://yfk.csb.gov.tr/aylik-guncel-rayic-ve-birim-fiyat-listeleri-113351" target="_blank" style={{color:"#60a5fa"}}>aylık liste sayfasından</a> ilgili ayı indir<br/>
          3. PDF'i aşağıya yükle → Claude otomatik parse eder → Onayla
        </div>
      </div>

      {/* Ayarlar */}
      {(durum === "bos" || durum === "hazir" || durum === "hata") && (
        <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
          <div style={{flex:1,minWidth:120}}>
            <label style={S.formEtiket}>Yıl</label>
            <select style={{...S.select,width:"100%"}} value={yil} onChange={e=>setYil(e.target.value)}>
              {["2026","2025","2024"].map(y=><option key={y}>{y}</option>)}
            </select>
          </div>
          <div style={{flex:1,minWidth:140}}>
            <label style={S.formEtiket}>Ay (aylık güncelleme ise)</label>
            <select style={{...S.select,width:"100%"}} value={ay} onChange={e=>setAy(e.target.value)}>
              <option value="">Yıllık liste</option>
              {AYLAR_TR.map(a=><option key={a}>{a}</option>)}
            </select>
          </div>
          <div style={{flex:2,minWidth:180}}>
            <label style={S.formEtiket}>Liste Tipi</label>
            <select style={{...S.select,width:"100%"}} value={listeTipi} onChange={e=>setListeTipi(e.target.value)}>
              {["İnşaat Birim Fiyat","Mekanik Tesisat Birim Fiyat","Elektrik Tesisat Birim Fiyat","İnşaat Rayiç","Mekanik Rayiç","Elektrik Rayiç"].map(t=><option key={t}>{t}</option>)}
            </select>
          </div>
        </div>
      )}

      {/* Yükleme alanı */}
      {(durum === "bos" || durum === "hata") && (
        <div
          style={{border:`2px dashed ${surukle?"#3b82f6":"#1e293b"}`,borderRadius:12,padding:"48px 24px",
            textAlign:"center",backgroundColor:surukle?"#0d1f3c":"#111827",transition:"all 0.2s",cursor:"pointer"}}
          onDragOver={e=>{e.preventDefault();setSurukle(true)}}
          onDragLeave={()=>setSurukle(false)}
          onDrop={e=>{e.preventDefault();setSurukle(false);dosyaOku(e.dataTransfer.files[0])}}
          onClick={()=>document.getElementById("pdfInput").click()}>
          <input id="pdfInput" type="file" accept=".pdf" style={{display:"none"}} onChange={e=>dosyaOku(e.target.files[0])}/>
          <div style={{fontSize:48,marginBottom:12}}>📄</div>
          <div style={{fontSize:15,fontWeight:700,color:"#e2e8f0",marginBottom:6}}>PDF'i buraya sürükle veya tıkla</div>
          <div style={{fontSize:12,color:"#475569"}}>Çevre ve Şehircilik Bakanlığı birim fiyat veya rayiç listesi</div>
          {hata && <div style={{marginTop:12,color:"#ef4444",fontSize:12,fontWeight:600}}>{hata}</div>}
        </div>
      )}

      {/* Dosya hazır */}
      {durum === "hazir" && (
        <div style={{backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:10,padding:20}}>
          <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16}}>
            <span style={{fontSize:32}}>📄</span>
            <div>
              <div style={{fontSize:14,fontWeight:700,color:"#e2e8f0"}}>{dosyaAdi}</div>
              <div style={{fontSize:12,color:"#475569",marginTop:2}}>{yil} {ay||"Yıllık"} — {listeTipi}</div>
            </div>
            <button style={{marginLeft:"auto",...S.ikinciBtn,padding:"6px 12px",fontSize:12}} onClick={()=>{setDurum("bos");setPdfBase64(null);}}>✕ Kaldır</button>
          </div>
          <button style={{...S.birincilBtn,width:"100%",padding:"12px 0",fontSize:14,backgroundColor:"#052e16",borderColor:"#16a34a",color:"#4ade80"}} onClick={pdfAnaliz}>
            🤖 Claude ile Analiz Et
          </button>
        </div>
      )}

      {/* Analiz sürüyor */}
      {durum === "analiz" && (
        <div style={{backgroundColor:"#111827",border:"1px solid #1e293b",borderRadius:12,padding:40,textAlign:"center"}}>
          <div style={{fontSize:48,marginBottom:16,animation:"pulse 1.5s infinite"}}>🤖</div>
          <div style={{fontSize:16,fontWeight:700,color:"#e2e8f0",marginBottom:8}}>PDF Analiz Ediliyor...</div>
          <div style={{fontSize:13,color:"#475569",marginBottom:20}}>Claude birim fiyat listesini okuyor ve pozları çıkarıyor</div>
          <div style={{display:"flex",gap:6,justifyContent:"center"}}>
            {[0,1,2].map(i=>(
              <div key={i} style={{width:8,height:8,borderRadius:"50%",backgroundColor:"#3b82f6",
                animation:`bounce 1s ${i*0.2}s infinite`}}/>
            ))}
          </div>
          <style>{`@keyframes bounce{0%,100%{transform:translateY(0)}50%{transform:translateY(-8px)}}`}</style>
        </div>
      )}

      {/* Onay */}
      {durum === "onay" && analizSonucu && (
        <div>
          <div style={{backgroundColor:"#052e16",border:"1px solid #16a34a",borderRadius:10,padding:"14px 18px",marginBottom:16,display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:24}}>✅</span>
            <div>
              <div style={{fontSize:14,fontWeight:700,color:"#4ade80"}}>Analiz Tamamlandı!</div>
              <div style={{fontSize:12,color:"#86efac",marginTop:2}}>
                {analizSonucu.pozlar.length} poz çıkarıldı — {analizSonucu.yil || yil} {analizSonucu.ay || ay || "Yıllık"} {listeTipi}
              </div>
            </div>
          </div>

          {/* Önizleme tablosu */}
          <div style={{...S.tablo,marginBottom:16,maxHeight:320,overflow:"auto"}}>
            <div style={S.tabloBaslik}>
              <span style={{flex:1}}>Poz No</span>
              <span style={{flex:3}}>Tarif</span>
              <span style={{flex:0.8,textAlign:"center"}}>Birim</span>
              <span style={{flex:1.2,textAlign:"right"}}>Birim Fiyat (₺)</span>
              <span style={{flex:1}}>Kategori</span>
            </div>
            {analizSonucu.pozlar.slice(0,50).map((p,i)=>(
              <div key={i} style={S.tabloSatir}>
                <span style={{flex:1,fontFamily:"monospace",fontSize:11,color:"#6366f1"}}>{p.no}</span>
                <span style={{flex:3,fontSize:11,color:"#cbd5e1"}}>{p.tanim}</span>
                <span style={{flex:0.8,textAlign:"center",fontSize:11,color:"#94a3b8"}}>{p.birim}</span>
                <span style={{flex:1.2,textAlign:"right",fontSize:12,color:"#10b981",fontWeight:700}}>₺{fmt(p.birimFiyat)}</span>
                <span style={{flex:1}}>
                  <span style={{fontSize:9,padding:"2px 6px",borderRadius:10,backgroundColor:(KAT_RENK[p.kategori]||"#475569")+"22",color:KAT_RENK[p.kategori]||"#475569"}}>{p.kategori}</span>
                </span>
              </div>
            ))}
            {analizSonucu.pozlar.length > 50 && (
              <div style={{padding:"10px 16px",fontSize:12,color:"#475569",textAlign:"center",borderTop:"1px solid #1e293b"}}>
                ... ve {analizSonucu.pozlar.length - 50} poz daha
              </div>
            )}
          </div>

          {/* Uygulama seçenekleri */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <button
              style={{...S.birincilBtn,padding:"14px 0",backgroundColor:"#052e16",borderColor:"#16a34a",color:"#4ade80",fontSize:13}}
              onClick={()=>pozlariUygula("yenile")}>
              🔄 Kütüphaneyi Tamamen Yenile
              <div style={{fontSize:10,color:"#86efac",fontWeight:400,marginTop:3}}>Mevcut {mevcutPozSayisi} poz silinir</div>
            </button>
            <button
              style={{...S.birincilBtn,padding:"14px 0",fontSize:13}}
              onClick={()=>pozlariUygula("ekle")}>
              ➕ Mevcut Kütüphaneye Ekle / Güncelle
              <div style={{fontSize:10,color:"#93c5fd",fontWeight:400,marginTop:3}}>Fiyatlar güncellenir, yeni pozlar eklenir</div>
            </button>
          </div>
          <button style={{...S.ikinciBtn,marginTop:8,width:"100%"}} onClick={()=>{setDurum("hazir");setAnalizSonucu(null);}}>← Geri Dön</button>
        </div>
      )}

      {/* Başarı */}
      {durum === "tamam" && (
        <div style={{textAlign:"center",padding:60}}>
          <div style={{fontSize:64,marginBottom:16}}>🎉</div>
          <div style={{fontSize:20,fontWeight:800,color:"#4ade80",marginBottom:8}}>Poz Kütüphanesi Güncellendi!</div>
          <div style={{fontSize:13,color:"#475569",marginBottom:24}}>
            {analizSonucu?.pozlar?.length} poz başarıyla yüklendi
          </div>
          <div style={{display:"flex",gap:12,justifyContent:"center"}}>
            <button style={{...S.birincilBtn}} onClick={()=>setAnaEkran("kutuphane")}>📚 Kütüphaneye Git</button>
            <button style={{...S.ikinciBtn}} onClick={()=>{setDurum("bos");setPdfBase64(null);setAnalizSonucu(null);}}>Başka PDF Yükle</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Hakediş Cetveli ───────────────────────────────────────────────
