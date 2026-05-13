import { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

// ── Sabitler ──────────────────────────────────────────────────────
const IDARELER = ["Tümü","KGM","DSİ","MEB","Çevre ve Şehircilik","Sağlık Bakanlığı","Belediye"];
const KATEGORILER = ["Tümü","Beton İşleri","Demir İşleri","Kaba İnşaat","İnce İnşaat","Tesisat","Elektrik","Zemin İşleri","Nakliye","Diğer"];

const POZLAR_BASLANGIC = [
  {id:100,no:"15.100.1001",tanim:"1 ton her cins çimento ve kirecin taşıtlara yükleme, boşaltma ve istifi",birim:"ton",birimFiyat:256.25,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:101,no:"15.100.1002",tanim:"",birim:"m³",birimFiyat:39.66,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:102,no:"15.100.1003",tanim:"1 m³ her nevi taşın taşıtlara yükleme boşaltma ve figüresi",birim:"m³",birimFiyat:43.44,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:103,no:"15.100.1004",tanim:"",birim:"ton",birimFiyat:175.31,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:104,no:"15.100.1005",tanim:"1 ton çelik borunun taşıtlara yükleme, boşaltma ve istifi",birim:"ton",birimFiyat:350.61,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:105,no:"15.100.1006",tanim:"",birim:"ton",birimFiyat:525.93,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:106,no:"15.100.1007",tanim:"",birim:"adet",birimFiyat:232.58,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:107,no:"15.100.1008",tanim:"",birim:"m³",birimFiyat:66.45,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:108,no:"15.105.1001",tanim:"Kazı alanı içine rastlayan fundaların gerekli el aletleri kullanarak kesilmesi ve temizlenmesi",birim:"m²",birimFiyat:6406.25,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:109,no:"15.105.1002",tanim:"Kazı ve dolgu alanında makine ile temizleme ve sökme işi yapılması",birim:"m²",birimFiyat:648.96,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:110,no:"15.105.1101",tanim:"El ile ağaç kesilmesi ve sökme işi, çapı 5-10 cm (10 cm dahil) beher ağaç için",birim:"adet",birimFiyat:128.13,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:111,no:"15.105.1102",tanim:"Ağaç kesilmesi ve sökülmesi",birim:"adet",birimFiyat:256.25,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:112,no:"15.105.1103",tanim:"El ile ağaç kesilmesi ve sökme işi, çapı 21-30 cm (30 cm dahil) beher ağaç için",birim:"adet",birimFiyat:512.5,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:113,no:"15.105.1104",tanim:"El ile ağaç kesilmesi ve sökme işi, çapı 31-40 cm (40 cm dahil) beher ağaç için",birim:"adet",birimFiyat:768.75,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:114,no:"15.105.1105",tanim:"El ile ağaç kesilmesi ve sökme işi, çapı 41-50 cm (50 cm dahil) beher ağaç için",birim:"adet",birimFiyat:1025.0,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:115,no:"15.105.1106",tanim:"El ile ağaç kesilmesi ve sökme işi, çapı 51-60 cm (60 cm dahil) beher ağaç için",birim:"adet",birimFiyat:1537.5,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:116,no:"15.105.1107",tanim:"El ile ağaç kesilmesi ve sökme işi, çapı 61-70 cm (70 cm dahil) beher ağaç için",birim:"adet",birimFiyat:2306.25,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:117,no:"15.105.1108",tanim:"El ile ağaç kesilmesi ve sökme işi, çapı 71-80 cm (80 cm dahil) beher ağaç için",birim:"adet",birimFiyat:3075.0,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:118,no:"15.105.1109",tanim:"El ile ağaç kesilmesi ve sökme işi, çapı 81 cm den büyük olan beher ağaç için",birim:"adet",birimFiyat:5125.0,idare:"Çevre ve Şehircilik",kategori:"Nakliye",kaynak:"2026"},
  {id:119,no:"15.106.1001",tanim:"Patlayıcı Madde Kullanmadan Harçsız Kagir İnşaatın Yıkılması",birim:"m³",birimFiyat:212.71,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:120,no:"15.106.1002",tanim:"Patlayıcı Madde Kullanmadan Kireç ve Melez Harçlı Kagir İnşaatın Yıkılması",birim:"m³",birimFiyat:270.03,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:121,no:"15.106.1003",tanim:"Patlayıcı Madde Kullanmadan Çimento Harçlı Kargir ve Horasan İnşaatın Yıkılması",birim:"m³",birimFiyat:365.55,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:122,no:"15.106.1004",tanim:"Patlayıcı Madde Kullanmadan Demirli ve Demirsiz Beton İnşaatın Yıkılması",birim:"m³",birimFiyat:671.24,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:123,no:"15.106.1005",tanim:"İnsan Gücü ile (Kompresör vasıtasıyla) Her Cins Harçlı Kargir İnşaatın Yıkılması",birim:"m³",birimFiyat:996.43,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:124,no:"15.106.1006",tanim:"İnsan Gücü ile (Kompresör Vasıtasıyla) Demirli ve Demirsiz Beton İnşaatın Yıkılması",birim:"m³",birimFiyat:1910.55,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:125,no:"15.106.1007",tanim:"Patlayıcı Madde Kullanarak Her Cins Harçlı Kargir ve Horasan İnşaatın Yıkılması",birim:"m³",birimFiyat:1572.48,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:126,no:"15.106.1008",tanim:"Patlayıcı Madde Kullanarak Demirli ve Demirsiz Beton İnşaatın Yıkılması",birim:"m³",birimFiyat:2367.19,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:127,no:"15.106.1009",tanim:"Demirli Beton İnşaatın Yıkılmasından Çıkan Demirlerin İnsan Gücü ile Ayrılması",birim:"kg",birimFiyat:12.81,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:128,no:"15.106.1010",tanim:"Demirli Beton İnşaatın Yıkılmasından Çıkan Demirlerin Makine Gücü ile Ayrılması",birim:"kg",birimFiyat:6.83,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:129,no:"15.106.1011",tanim:"Her Türlü Yıkımdan Çıkan Taş ve Tuğlanın İnsan Gücü ile Ayrılması",birim:"m³",birimFiyat:768.75,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:130,no:"15.106.1101",tanim:"Doğal parke taşı, beton plak, adi kaldırım ve blokaj sökülmesi",birim:"m²",birimFiyat:256.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:131,no:"15.106.1102",tanim:"Kırma taş, şose ve asfalt sökülmesi",birim:"m³",birimFiyat:1409.38,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:132,no:"15.106.1103",tanim:"Her türlü bordür sökülmesi",birim:"m",birimFiyat:64.06,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:133,no:"15.106.1104",tanim:"Her türlü iç sıva sökülmesi",birim:"m²",birimFiyat:153.75,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:134,no:"15.106.1105",tanim:"Her türlü dış sıva sökülmesi",birim:"m²",birimFiyat:281.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:135,no:"15.106.1106",tanim:"Her türlü ahşap çatı sökülmesi",birim:"m²",birimFiyat:514.06,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:136,no:"15.106.1107",tanim:"Her türlü kiremit örtülü çatılarda kiremit aktarılması",birim:"m²",birimFiyat:202.44,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:137,no:"15.106.1108",tanim:"Her türlü kiremit çatı örtüsü sökülmesi, toplanması, temizlenmesi, istif edilmesi",birim:"m²",birimFiyat:115.63,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:138,no:"15.106.1109",tanim:"Galvanizli sac, alüminyum, cam elyaf takviyeli vb. çatı örtüsü sökülmesi",birim:"m²",birimFiyat:180.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:139,no:"15.106.1110",tanim:"Kenetli sac ve bakır çatı örtüsü sökülmesi",birim:"m²",birimFiyat:270.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:140,no:"15.106.1111",tanim:"Çatı örtüsü altındaki ahşap kaplama tahtası sökülmesi",birim:"m²",birimFiyat:200.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:141,no:"15.106.1112",tanim:"Çatı örtüleri altındaki su yalıtım örtüleri ve bitümlü karton sökülmesi",birim:"m²",birimFiyat:50.31,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:142,no:"15.106.1113",tanim:"Her türlü, sac, pvc, çinko vb, yağmur oluğu ve borusu sökülmesi",birim:"m",birimFiyat:94.69,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:143,no:"15.106.1114",tanim:"",birim:"m",birimFiyat:186.56,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:144,no:"15.106.1115",tanim:"Çinko ve sacdan her türlü atika duvar arkası çatı deresi (gizli dere) sökülmesi",birim:"m",birimFiyat:308.75,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:145,no:"15.106.1116",tanim:"Her kalınlıkta mermer, traverten, terraza karo ve andezit kaplama sökülmesi",birim:"m²",birimFiyat:266.25,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:146,no:"15.106.1117",tanim:"Seramik, fayans vb. kaplama sökülmesi",birim:"m²",birimFiyat:303.75,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:147,no:"15.106.1118",tanim:"Dökme mozayik sökülmesi",birim:"m²",birimFiyat:631.25,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:148,no:"15.106.1119",tanim:"PVC vb. döşeme kaplaması sökülmesi",birim:"m²",birimFiyat:150.88,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:149,no:"15.106.1120",tanim:"Ahşap parke ve ahşap döşeme kaplaması sökülmesi",birim:"m²",birimFiyat:320.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:150,no:"15.106.1121",tanim:"Ahşap ve PVC vb. süpürgelik sökülmesi",birim:"m",birimFiyat:62.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:151,no:"15.106.1122",tanim:"Ahşap lambiri vb. duvar kaplaması sökülmesi",birim:"m²",birimFiyat:400.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:152,no:"15.106.1123",tanim:"Mermer, traverten ve andezit plaktan denizlik, parapet ve harpuşta sökülmesi",birim:"m²",birimFiyat:379.69,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:153,no:"15.106.1124",tanim:"Dökme mozayikten denizlik, parapet ve harpuşta sökülmesi",birim:"m²",birimFiyat:631.25,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:154,no:"15.106.1125",tanim:"Duvar, döşeme ve tavanda sac, alüminyum, ahşap vb, dilatasyon fugası sökülmesi",birim:"m",birimFiyat:266.25,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:155,no:"15.106.1126",tanim:"Her türlü ahşap dolap sökülmesi",birim:"m²",birimFiyat:400.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:156,no:"15.106.1127",tanim:"Tezgah üstü mermeri sökülmesi",birim:"m²",birimFiyat:266.25,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:157,no:"15.106.1128",tanim:"Ahşap asma tavan sökülmesi",birim:"m²",birimFiyat:428.75,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:158,no:"15.106.1129",tanim:"Alüminyum, sac, alçıpanel, taşyünü, camyünü vb. asma tavan sökülmesi",birim:"m²",birimFiyat:310.63,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:159,no:"15.106.1130",tanim:"Alüminyum ve PVC den yapılan her türlü kapı ve pencere doğramasının sökülmesi",birim:"m²",birimFiyat:221.88,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:160,no:"15.106.1131",tanim:"Her türlü ahşap kapı kasası, kapı kanadı, pencere ve camekan sökülmesi",birim:"m²",birimFiyat:257.81,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:161,no:"15.106.1132",tanim:"",birim:"kg",birimFiyat:64.96,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:162,no:"15.115.1001",tanim:"",birim:"m³",birimFiyat:512.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:163,no:"15.115.1002",tanim:"",birim:"m³",birimFiyat:666.25,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:164,no:"15.115.1003",tanim:"parçalanıp el ile atılabilen 0,100 m3 e kadar büyüklükteki her cins blok taşlar, kazı güçlüğü",birim:"m³",birimFiyat:832.81,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:165,no:"15.115.1004",tanim:"marn ve kil, 0,100-0,400 m³ büyüklükte parçalanıp el ile atılabilen her cins blok taşlar ve",birim:"m³",birimFiyat:960.94,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:166,no:"15.115.1005",tanim:"kalker, marnlı kalker, marn, şist, gre, gevşek konglomera, alçı taşı, volkanik tüfler (bazalt",birim:"m³",birimFiyat:768.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:167,no:"15.115.1006",tanim:"halinde sert gre, betonlaşmış konglomera, kesif kalker, mermer, ayrışmamış serpantin,",birim:"m³",birimFiyat:936.91,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:168,no:"15.115.1007",tanim:"ve benzeri, bazalt, profir, kuvars, 0,400 m³ den büyük aynı cins blok taşlar ve benzeri",birim:"m³",birimFiyat:1137.66,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:169,no:"15.115.1008",tanim:"kalker, marnlı kalker, marn, şist, gre, gevşek konglomera, alçı taşı, volkanik tüfler (bazalt",birim:"m³",birimFiyat:1306.09,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:170,no:"15.115.1009",tanim:"konglomera, kesif kalker, mermer, ayrışmamış serpantin, andezit, trakit bazalt tüfleri ve",birim:"m³",birimFiyat:1776.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:171,no:"15.115.1010",tanim:"granit ve benzeri, bazalt, profir, kuvars, 0,400 m³ den büyük aynı cins blok taşlar ve benzeri",birim:"m³",birimFiyat:2679.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:172,no:"15.115.1011",tanim:"",birim:"m³",birimFiyat:1537.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:173,no:"15.115.1201",tanim:"toprak, gevşek silt, kum, kil, siltli, kumlu ve gevşek kil, killi kum ve çakıl, kürekle atılabilen",birim:"m³",birimFiyat:960.94,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:174,no:"15.115.1202",tanim:"toprak, gevşek silt, kum, kil, siltli, kumlu ve gevşek kil, killi kum ve çakıl, kürekle atılabilen",birim:"m³",birimFiyat:1057.04,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:175,no:"15.115.1203",tanim:"",birim:"m³",birimFiyat:1383.75,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:176,no:"15.115.1204",tanim:"",birim:"m³",birimFiyat:1522.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:177,no:"15.115.1206",tanim:"",birim:"m³",birimFiyat:1877.09,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:178,no:"15.115.1207",tanim:"",birim:"m³",birimFiyat:1876.9,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:179,no:"15.115.1208",tanim:"",birim:"m³",birimFiyat:2064.59,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:180,no:"15.115.1209",tanim:"",birim:"m³",birimFiyat:2347.16,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:181,no:"15.115.1210",tanim:"",birim:"m³",birimFiyat:2581.88,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:182,no:"15.115.1211",tanim:"derin kazı yapılması (ayrışmamış granit ve benzeri, bazalt, profir, kuvars, 0,400 m³ den",birim:"m³",birimFiyat:3287.7,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:183,no:"15.115.1212",tanim:"kazı yapılması (ayrışmamış granit ve benzeri, bazalt, profir, kuvars, 0,400 m³ den büyük",birim:"m³",birimFiyat:3616.48,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:184,no:"15.115.1213",tanim:"",birim:"m³",birimFiyat:263.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:185,no:"15.115.1215",tanim:"",birim:"m³",birimFiyat:2319.06,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:186,no:"15.115.1216",tanim:"",birim:"m³",birimFiyat:2550.98,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:187,no:"15.120.1001",tanim:"Makine ile yumuşak ve sert toprak kazılması (Serbest kazı)",birim:"m³",birimFiyat:62.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:188,no:"15.120.1002",tanim:"Makine ile yumuşak ve sert küskülük kazılması (serbest kazı)",birim:"m³",birimFiyat:83.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:189,no:"15.120.1003",tanim:"Makine ile batak ve balçık kazılması (serbest kazı)",birim:"m³",birimFiyat:122.64,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:190,no:"15.120.1004",tanim:"Makine ile patlayıcı madde kullanılarak yumuşak kaya kazılması (Serbest kazı)",birim:"m³",birimFiyat:187.94,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:191,no:"15.120.1005",tanim:"Makine ile patlayıcı madde kullanmadan yumuşak kaya kazılması (Serbest kazı)",birim:"m³",birimFiyat:138.39,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:192,no:"15.120.1006",tanim:"Makine ile patlayıcı madde kullanılarak sert kaya kazılması (Serbest kazı)",birim:"m³",birimFiyat:250.45,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:193,no:"15.120.1007",tanim:"Makine ile patlayıcı madde kullanmadan sert kaya kazılması (Serbest kazı)",birim:"m³",birimFiyat:318.26,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:194,no:"15.120.1008",tanim:"Makine ile patlayıcı madde kullanılarak çok sert kaya kazılması (Serbest kazı)",birim:"m³",birimFiyat:328.28,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:195,no:"15.120.1009",tanim:"Makine ile patlayıcı madde kullanmadan çok sert kaya kazılması (Serbest kazı)",birim:"m³",birimFiyat:430.3,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:196,no:"15.120.1101",tanim:"Makine ile her derinlik ve her genişlikte yumuşak ve sert toprak kazılması (Derin kazı)",birim:"m³",birimFiyat:71.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:197,no:"15.120.1102",tanim:"Makine ile her derinlik ve her genişlikte yumuşak ve sert küskülük kazılması (Derin kazı)",birim:"m³",birimFiyat:105.71,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:198,no:"15.120.1103",tanim:"Makine ile her derinlik ve her genişlikte batak ve balçık kazılması (Derin kazı)",birim:"m³",birimFiyat:167.64,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:199,no:"15.120.1105",tanim:"",birim:"m³",birimFiyat:172.46,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:200,no:"15.120.1107",tanim:"",birim:"m³",birimFiyat:413.28,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:201,no:"15.120.1109",tanim:"",birim:"m³",birimFiyat:517.43,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:202,no:"15.125.1001",tanim:"Kum temin edilerek, el ile serme, sulama ve sıkıştırma yapılması",birim:"m³",birimFiyat:588.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:203,no:"15.125.1002",tanim:"Çakıl temin edilerek, el ile serme, sulama ve sıkıştırma yapılması",birim:"m³",birimFiyat:585.63,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:204,no:"15.125.1003",tanim:"Kum temin edilerek, makine ile serme, sulama ve sıkıştırma yapılması",birim:"m³",birimFiyat:257.78,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:205,no:"15.125.1004",tanim:"Çakıl temin edilerek, makine ile serme, sulama ve sıkıştırma yapılması",birim:"m³",birimFiyat:255.28,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:206,no:"15.125.1005",tanim:"Kum temin edilerek, drenaj yapılması",birim:"m³",birimFiyat:937.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:207,no:"15.125.1006",tanim:"Çakıl temin edilerek, drenaj yapılması",birim:"m³",birimFiyat:937.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:208,no:"15.125.1007",tanim:"32mm'ye kadar kırmataş temin edilerek, el ile serme, sulama ve sıkıştırma yapılması",birim:"m³",birimFiyat:1023.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:209,no:"15.125.1008",tanim:"32mm'ye kadar kırmataş temin edilerek, makine ile serme, sulama ve sıkıştırma yapılması",birim:"m³",birimFiyat:692.78,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:210,no:"15.125.1009",tanim:"63 mm'ye kadar kırmataş temin edilerek, el ile serme, sulama ve sıkıştırma yapılması",birim:"m³",birimFiyat:948.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:211,no:"15.125.1010",tanim:"63mm'ye kadar kırmataş temin edilerek, makine ile serme, sulama ve sıkıştırma yapılması",birim:"m³",birimFiyat:617.78,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:212,no:"15.125.1011",tanim:"Hafif agrega (elenmiş kömür curufu) ile dolgu yapılması",birim:"m³",birimFiyat:160.63,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:213,no:"15.130.1001",tanim:"Kazılara tam ahşap kaplamalı iksa yapılması",birim:"m²",birimFiyat:1135.63,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:214,no:"15.130.1002",tanim:"Kazılara tam ahşap kaplamalı iksa yapılması",birim:"m²",birimFiyat:1135.63,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:215,no:"15.130.1003",tanim:"Kazılara sık aralıklı ahşap iksa yapılması",birim:"m²",birimFiyat:794.94,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:216,no:"15.130.1004",tanim:"Kazılara aralıklı ahşap kaplamalı iksa yapılması",birim:"m²",birimFiyat:567.81,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:217,no:"15.135.1001",tanim:"",birim:"m",birimFiyat:1153.94,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:218,no:"15.135.1002",tanim:"",birim:"m",birimFiyat:1338.46,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:219,no:"15.135.1003",tanim:"",birim:"m",birimFiyat:1336.06,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:220,no:"15.135.1004",tanim:"",birim:"m",birimFiyat:1544.88,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:221,no:"15.135.1005",tanim:"",birim:"m",birimFiyat:1836.53,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:222,no:"15.140.1001",tanim:"",birim:"m",birimFiyat:1567.28,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:223,no:"15.140.1002",tanim:"",birim:"m",birimFiyat:2052.3,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:224,no:"15.140.1003",tanim:"",birim:"m",birimFiyat:3441.38,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:225,no:"15.140.1004",tanim:"",birim:"m",birimFiyat:3714.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:226,no:"15.140.1005",tanim:"",birim:"m",birimFiyat:4529.16,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:227,no:"15.140.1006",tanim:"",birim:"m",birimFiyat:4984.36,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:228,no:"15.140.1007",tanim:"",birim:"m",birimFiyat:6759.66,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:229,no:"15.140.1008",tanim:"",birim:"m",birimFiyat:7562.86,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:230,no:"15.140.1009",tanim:"",birim:"m",birimFiyat:9327.38,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:231,no:"15.140.1010",tanim:"",birim:"m",birimFiyat:10474.83,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:232,no:"15.140.1011",tanim:"",birim:"m",birimFiyat:15545.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:233,no:"15.140.1012",tanim:"",birim:"m",birimFiyat:17840.24,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:234,no:"15.140.1101",tanim:"",birim:"m",birimFiyat:1577.28,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:235,no:"15.140.1102",tanim:"",birim:"m",birimFiyat:2073.55,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:236,no:"15.140.1103",tanim:"",birim:"m",birimFiyat:3486.38,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:237,no:"15.140.1104",tanim:"",birim:"m",birimFiyat:3759.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:238,no:"15.140.1105",tanim:"",birim:"m",birimFiyat:4597.91,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:239,no:"15.140.1106",tanim:"",birim:"m",birimFiyat:5053.11,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:240,no:"15.140.1107",tanim:"",birim:"m",birimFiyat:6867.16,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:241,no:"15.140.1108",tanim:"",birim:"m",birimFiyat:7670.36,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:242,no:"15.140.1109",tanim:"",birim:"m",birimFiyat:9482.38,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:243,no:"15.140.1110",tanim:"",birim:"m",birimFiyat:10629.83,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:244,no:"15.140.1111",tanim:"",birim:"m",birimFiyat:15839.1,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:245,no:"15.140.1112",tanim:"",birim:"m",birimFiyat:18133.99,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:246,no:"15.140.1201",tanim:"",birim:"m",birimFiyat:1587.28,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:247,no:"15.140.1202",tanim:"",birim:"m",birimFiyat:2094.8,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:248,no:"15.140.1203",tanim:"",birim:"m",birimFiyat:3531.38,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:249,no:"15.140.1204",tanim:"",birim:"m",birimFiyat:3804.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:250,no:"15.140.1205",tanim:"",birim:"m",birimFiyat:4666.66,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:251,no:"15.140.1206",tanim:"",birim:"m",birimFiyat:5121.86,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:252,no:"15.140.1207",tanim:"",birim:"m",birimFiyat:6974.66,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:253,no:"15.140.1208",tanim:"",birim:"m",birimFiyat:7777.86,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:254,no:"15.140.1209",tanim:"",birim:"m",birimFiyat:9637.38,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:255,no:"15.140.1210",tanim:"",birim:"m",birimFiyat:10784.83,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:256,no:"15.140.1211",tanim:"",birim:"m",birimFiyat:16132.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:257,no:"15.140.1212",tanim:"",birim:"m",birimFiyat:18427.74,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:258,no:"15.150.1001",tanim:"",birim:"m³",birimFiyat:2964.43,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:259,no:"15.150.1002",tanim:"",birim:"m³",birimFiyat:3089.43,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:260,no:"15.150.1003",tanim:"",birim:"m³",birimFiyat:3274.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:261,no:"15.150.1004",tanim:"",birim:"m³",birimFiyat:3399.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:262,no:"15.150.1005",tanim:"",birim:"m³",birimFiyat:3524.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:263,no:"15.150.1006",tanim:"",birim:"m³",birimFiyat:3649.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:264,no:"15.150.1007",tanim:"",birim:"m³",birimFiyat:3836.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:265,no:"15.150.1008",tanim:"",birim:"m³",birimFiyat:3961.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:266,no:"15.150.1009",tanim:"",birim:"m³",birimFiyat:4086.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:267,no:"15.150.1010",tanim:"",birim:"m³",birimFiyat:4211.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:268,no:"15.150.1101",tanim:"",birim:"m³",birimFiyat:3495.68,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:269,no:"15.150.1102",tanim:"",birim:"m³",birimFiyat:3620.68,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:270,no:"15.150.1103",tanim:"",birim:"m³",birimFiyat:3836.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:271,no:"15.150.1104",tanim:"",birim:"m³",birimFiyat:3961.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:272,no:"15.150.1105",tanim:"",birim:"m³",birimFiyat:4086.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:273,no:"15.150.1106",tanim:"",birim:"m³",birimFiyat:4461.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:274,no:"15.150.1107",tanim:"",birim:"m³",birimFiyat:4774.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:275,no:"15.150.1108",tanim:"",birim:"m³",birimFiyat:5149.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:276,no:"15.150.1109",tanim:"",birim:"m³",birimFiyat:5461.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:277,no:"15.150.1110",tanim:"",birim:"m³",birimFiyat:5836.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:278,no:"15.150.5101",tanim:"Kayar Kalıplı Beton Finişeri ile Yol Betonu Serilmesi (C30/37 Beton Dayanım Sınıfında)",birim:"m³",birimFiyat:5357.55,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:279,no:"15.150.5102",tanim:"Kayar Kalıplı Beton Finişeri ile Yol Betonu Serilmesi (C35/40 Beton Dayanım Sınıfında)",birim:"m³",birimFiyat:5545.05,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:280,no:"15.150.5103",tanim:"Kayar Kalıplı Beton Finişeri ile Yol Betonu Serilmesi (C40/50 Beton Dayanım Sınıfında)",birim:"m³",birimFiyat:5670.05,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:281,no:"15.150.5111",tanim:"Beton Yol Yüzeyinin Elle Kürlenmesi, Perdahlanması ve Pürüzlendirilmesi",birim:"m²",birimFiyat:90.63,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:282,no:"15.150.5112",tanim:"",birim:"m²",birimFiyat:139.31,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:283,no:"15.155.1007",tanim:"",birim:"m²",birimFiyat:1966.9,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:284,no:"15.155.1008",tanim:"",birim:"m²",birimFiyat:2094.4,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:285,no:"15.155.1011",tanim:"",birim:"m²",birimFiyat:2110.09,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:286,no:"15.155.1012",tanim:"",birim:"m²",birimFiyat:2309.19,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:287,no:"15.155.1013",tanim:"",birim:"m²",birimFiyat:2575.54,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:288,no:"15.155.1014",tanim:"",birim:"m²",birimFiyat:2703.04,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:289,no:"15.155.1015",tanim:"",birim:"m²",birimFiyat:3081.03,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:290,no:"15.155.1016",tanim:"",birim:"m²",birimFiyat:3399.78,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:291,no:"15.160.1001",tanim:"",birim:"ton",birimFiyat:44531.25,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:292,no:"15.160.1002",tanim:"",birim:"ton",birimFiyat:43987.5,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:293,no:"15.160.1003",tanim:"",birim:"ton",birimFiyat:46482.43,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:294,no:"15.160.1004",tanim:"",birim:"ton",birimFiyat:44681.18,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:295,no:"15.160.1005",tanim:"",birim:"ton",birimFiyat:43749.93,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:296,no:"15.165.1001",tanim:"kullanılan münferit çatı aşıkları ve mertekleri, lentolar, hurdi döşemeler, köşe takviye",birim:"ton",birimFiyat:85414.38,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:297,no:"15.165.1002",tanim:"Profil demirlerinden çatı makası yapılması ve yerine konulması.",birim:"ton",birimFiyat:92835.66,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:298,no:"15.165.1003",tanim:"(yapı karkası, köprülerde profil demirlerinden kirişler, başlıklar, bağlantılar ve benzeri",birim:"ton",birimFiyat:85943.51,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:299,no:"15.180.1001",tanim:"Ahşaptan seri kalıp yapılması",birim:"m²",birimFiyat:331.63,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:300,no:"15.180.1002",tanim:"Ahşaptan düz yüzeyli beton ve betonarme kalıbı yapılması",birim:"m²",birimFiyat:873.03,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:301,no:"15.180.1004",tanim:"Sac ile eğri yüzeyli beton ve betonarme kalıbı yapılması",birim:"m²",birimFiyat:983.15,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:302,no:"15.185.1005",tanim:"Çelik borudan kalıp iskelesi yapılması (0,00-4,00 m arası)",birim:"m³",birimFiyat:127.75,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:303,no:"15.185.1006",tanim:"Çelik borudan kalıp iskelesi yapılması (4,01-6,00 m arası)",birim:"m³",birimFiyat:147.68,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:304,no:"15.185.1007",tanim:"Çelik borudan kalıp iskelesi yapılması (6,01-8,00m arası)",birim:"m³",birimFiyat:167.61,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:305,no:"15.185.1008",tanim:"Çelik borudan kalıp iskelesi yapılması 8,01-10,00m arası)",birim:"m³",birimFiyat:187.55,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:306,no:"15.185.1021",tanim:"",birim:"m²",birimFiyat:202.44,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:307,no:"15.185.1022",tanim:"",birim:"m²",birimFiyat:256.81,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:308,no:"15.185.1023",tanim:"",birim:"m²",birimFiyat:234.31,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:309,no:"15.185.1024",tanim:"",birim:"m²",birimFiyat:234.31,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:310,no:"15.185.1041",tanim:"Güvenlik Ağı Yapılması Sistem T (TS EN 1263-1)",birim:"m²",birimFiyat:38.6,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:311,no:"15.185.1042",tanim:"Güvenlik Ağı Yapılması (Sistem S ve Sistem U) TS EN 1263-1",birim:"m²",birimFiyat:38.6,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:312,no:"15.190.1001",tanim:"Bazalt agregalı (gri) yüzey sertlestirici ve kür uygulaması (taze betonda)",birim:"m²",birimFiyat:176.13,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:313,no:"15.190.1002",tanim:"Kuvars agregalı (gri) yüzey sertlestirici ve kür uygulaması (taze betonda)",birim:"m²",birimFiyat:179.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:314,no:"15.190.1003",tanim:"Kuvars-Korund agregalı (gri) yüzey sertlestirici ve kür uygulaması (taze betonda)",birim:"m²",birimFiyat:194.88,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:315,no:"15.190.1004",tanim:"Korund agregalı (gri) yüzey sertlestirici ve kür uygulaması (taze betonda)",birim:"m²",birimFiyat:207.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:316,no:"15.190.1005",tanim:"",birim:"m",birimFiyat:104.9,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:317,no:"15.190.1006",tanim:"Yeni beton yüzeylere kür yapılması (Saha Betonu)",birim:"m²",birimFiyat:36.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:318,no:"15.190.1007",tanim:"",birim:"m²",birimFiyat:197.71,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:319,no:"15.190.1008",tanim:"Silindir ile Sıkıştrılmış Beton Yollarda Parafinik Esaslı Kür Malzemesi ile Kür Yapılması",birim:"m²",birimFiyat:57.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:320,no:"15.190.1009",tanim:"Silindir ile Sıkıştrılmış Beton Yollarda Akrilik Esaslı Kür Malzemesi ile Kür Yapılması",birim:"m²",birimFiyat:60.75,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:321,no:"15.190.1011",tanim:"",birim:"m",birimFiyat:34.26,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:322,no:"15.190.1015",tanim:"bakteriostatik, iki bileşenli, poliüretan esaslı, şeffaf veya pigmentli, mat yüzey bitişli son kat",birim:"m²",birimFiyat:231.25,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:323,no:"15.190.1016",tanim:"uygulaması üzerine, anti-statik, iki bileşenli, poliüretan esaslı, mat, su bazlı ve düşük",birim:"m²",birimFiyat:460.63,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:324,no:"15.190.1017",tanim:"",birim:"m²",birimFiyat:214.25,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:325,no:"15.190.1018",tanim:"",birim:"m²",birimFiyat:320.0,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:326,no:"15.190.1019",tanim:"Alçı esaslı kendiliğinden yerleşen harç ile ortalama 2 mm kalınlıkta zemin tesviyesi yapılması",birim:"m²",birimFiyat:176.09,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:327,no:"15.195.1001",tanim:"",birim:"m",birimFiyat:738.99,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:328,no:"15.195.1002",tanim:"",birim:"m",birimFiyat:913.99,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:329,no:"15.195.1003",tanim:"",birim:"m",birimFiyat:1176.49,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:330,no:"15.195.1004",tanim:"",birim:"m",birimFiyat:5131.49,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:331,no:"15.200.1001",tanim:"esaslı drenaj ve koruma levhası temini ve yerine döşenmesi",birim:"m²",birimFiyat:110.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:332,no:"15.200.1002",tanim:"esaslı drenaj ve koruma levhası temini ve yerine döşenmesi",birim:"m²",birimFiyat:133.63,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:333,no:"15.200.1003",tanim:"esaslı drenaj ve koruma levhası temini ve yerine döşenmesi",birim:"m²",birimFiyat:158.56,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:334,no:"15.200.1004",tanim:"yerine döşenmesi",birim:"m²",birimFiyat:130.31,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:335,no:"15.200.1005",tanim:"yerine döşenmesi",birim:"m²",birimFiyat:153.94,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:336,no:"15.200.1006",tanim:"yerine döşenmesi",birim:"m²",birimFiyat:178.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:337,no:"15.204.1001",tanim:"Ø 100 mm anma çaplı, PVC esaslı koruge drenaj borusunun temini ve yerine döşenmesi",birim:"m",birimFiyat:56.0,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:338,no:"15.204.1002",tanim:"Ø 125 mm anma çaplı, PVC esaslı koruge drenaj borusunun temini ve yerine döşenmesi",birim:"m",birimFiyat:84.75,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:339,no:"15.204.1003",tanim:"Ø 160 mm anma çaplı, PVC esaslı koruge drenaj borusunun temini ve yerine döşenmesi",birim:"m",birimFiyat:122.25,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:340,no:"15.204.1004",tanim:"Ø 200 mm anma çaplı, PVC esaslı koruge drenaj borusunun temini ve yerine döşenmesi",birim:"m",birimFiyat:166.0,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:341,no:"15.210.1001",tanim:"Ocak taşı ile kuru duvar inşaat yapılması",birim:"m³",birimFiyat:1496.54,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:342,no:"15.210.1004",tanim:"Ocak taşı ile blokaj yapılması",birim:"m³",birimFiyat:1412.41,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:343,no:"15.210.1005",tanim:"Ocak taşı ile 200 dozlu çimento harçlı kargir inşaat yapılması",birim:"m³",birimFiyat:2384.78,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:344,no:"15.210.1006",tanim:"Ocaktan çaplanmış moloz taşı ile 200 dozlu çimento harçlı kargir inşaat yapılması",birim:"m³",birimFiyat:3911.66,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:345,no:"15.220.1011",tanim:"85 mm kalınlığında yatay delikli tuğla (190 x 85 x 190 mm) ile duvar yapılması",birim:"m²",birimFiyat:730.71,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:346,no:"15.220.1012",tanim:"100 mm kalınlığında yatay delikli tuğla (200 x 100 x 200 mm) ile duvar yapılması",birim:"m²",birimFiyat:753.39,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:347,no:"15.220.1013",tanim:"120 mm kalınlığında yatay delikli tuğla (250 x 120 x 200 mm) ile duvar yapılması",birim:"m²",birimFiyat:782.56,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:348,no:"15.220.1014",tanim:"135 mm kalınlığında yatay delikli tuğla (190 x 135 x 190 mm) ile duvar yapılması",birim:"m²",birimFiyat:808.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:349,no:"15.220.1015",tanim:"190 mm kalınlığında yatay delikli tuğla (190 x 190 x 135 mm) ile duvar yapılması",birim:"m²",birimFiyat:922.59,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:350,no:"15.220.1016",tanim:"200 mm kalınlığında yatay delikli tuğla (250 x 200 x 250 mm) ile duvar yapılması",birim:"m²",birimFiyat:993.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:351,no:"15.220.1017",tanim:"240 mm kalınlığında yatay delikli tuğla (235 x 240 x 135 mm) ile duvar yapılması",birim:"m²",birimFiyat:1110.73,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:352,no:"15.220.1018",tanim:"250 mm kalınlığında yatay delikli tuğla (240 x 250 x 190 mm) ile duvar yapılması",birim:"m²",birimFiyat:1119.96,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:353,no:"15.220.1121",tanim:"",birim:"m²",birimFiyat:902.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:354,no:"15.220.1122",tanim:"",birim:"m²",birimFiyat:993.16,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:355,no:"15.220.1123",tanim:"",birim:"m²",birimFiyat:1092.83,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:356,no:"15.220.1124",tanim:"",birim:"m²",birimFiyat:1143.7,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:357,no:"15.220.1125",tanim:"",birim:"m²",birimFiyat:1330.38,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:358,no:"15.220.1126",tanim:"",birim:"m²",birimFiyat:1367.31,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:359,no:"15.220.1127",tanim:"",birim:"m²",birimFiyat:1531.19,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:360,no:"15.220.1211",tanim:"",birim:"m²",birimFiyat:1044.44,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:361,no:"15.220.1212",tanim:"",birim:"m²",birimFiyat:1151.13,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:362,no:"15.220.1213",tanim:"",birim:"m²",birimFiyat:1315.29,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:363,no:"15.220.1214",tanim:"",birim:"m²",birimFiyat:1578.2,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:364,no:"15.220.1311",tanim:"90 mm kalınlığında düşey delikli dış cephe tuğlası (190 x 90 x 50 mm) ile duvar yapılması",birim:"m²",birimFiyat:2048.7,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:365,no:"15.220.1312",tanim:"102 mm kalınlığında düşey delikli dış cephe tuğlası (215 x 102 x 65 mm) ile duvar yapılması",birim:"m²",birimFiyat:2549.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:366,no:"15.220.1411",tanim:"190 mm kalınlığında düşey delikli tuğla (290 x 190 x 135 mm) ile duvar yapılması",birim:"m²",birimFiyat:1134.44,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:367,no:"15.220.1412",tanim:"290 mm kalınlığında düşey delikli tuğla (190 x 290 x 135 mm) ile duvar yapılması",birim:"m²",birimFiyat:1438.94,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:368,no:"15.220.1461",tanim:"90 mm kalınlığında dolu harman tuğlası (190 x 90 x 50 mm) ile duvar yapılması",birim:"m²",birimFiyat:1246.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:369,no:"15.220.1462",tanim:"90 mm kalınlığında delikli harman tuğlası (190 x 90 x 50 mm) ile duvar yapılması",birim:"m²",birimFiyat:1246.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:370,no:"15.220.1501",tanim:"",birim:"m²",birimFiyat:816.68,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:371,no:"15.220.1502",tanim:"",birim:"m²",birimFiyat:905.16,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:372,no:"15.220.1503",tanim:"",birim:"m²",birimFiyat:991.18,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:373,no:"15.220.1504",tanim:"",birim:"m²",birimFiyat:1085.39,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:374,no:"15.220.1505",tanim:"",birim:"m²",birimFiyat:1163.2,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:375,no:"15.220.1506",tanim:"",birim:"m²",birimFiyat:1257.43,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:376,no:"15.220.1507",tanim:"",birim:"m²",birimFiyat:1351.64,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:377,no:"15.220.1612",tanim:"12 cm - 13,5 cm kalınlıklar arasında teçhizatlı tuğla lentonun temini ve yerine konulması",birim:"m",birimFiyat:1284.74,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:378,no:"15.220.1613",tanim:"14,5 cm - 16 cm kalınlıklar arasında teçhizatlı tuğla lentonun temini ve yerine konulması",birim:"m",birimFiyat:1336.95,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:379,no:"15.220.1614",tanim:"18,5 cm - 20 cm kalınlıklar arasında teçhizatlı tuğla lentonun temini ve yerine konulması",birim:"m",birimFiyat:1502.96,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:380,no:"15.220.1615",tanim:"23,5 cm - 25 cm kalınlıklar arasında teçhizatlı tuğla lentonun temini ve yerine konulması",birim:"m",birimFiyat:1632.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:381,no:"15.225.1001",tanim:"",birim:"m²",birimFiyat:624.85,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:382,no:"15.225.1002",tanim:"",birim:"m²",birimFiyat:661.86,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:383,no:"15.225.1003",tanim:"",birim:"m²",birimFiyat:685.49,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:384,no:"15.225.1004",tanim:"",birim:"m²",birimFiyat:723.73,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:385,no:"15.225.1005",tanim:"",birim:"m²",birimFiyat:816.13,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:386,no:"15.225.1006",tanim:"",birim:"m²",birimFiyat:854.43,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:387,no:"15.225.1007",tanim:"",birim:"m²",birimFiyat:908.56,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:388,no:"15.225.1008",tanim:"",birim:"m²",birimFiyat:1001.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:389,no:"15.225.1009",tanim:"",birim:"m²",birimFiyat:1061.58,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:390,no:"15.225.1010",tanim:"",birim:"m²",birimFiyat:1106.28,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:391,no:"15.225.1011",tanim:"",birim:"m²",birimFiyat:1198.73,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:392,no:"15.225.1012",tanim:"",birim:"m²",birimFiyat:1291.1,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:393,no:"15.225.1013",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1386.73,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:394,no:"15.225.1014",tanim:"",birim:"m²",birimFiyat:1482.38,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:395,no:"15.225.1015",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1578.05,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:396,no:"15.225.1016",tanim:"",birim:"m²",birimFiyat:1673.84,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:397,no:"15.225.1051",tanim:"",birim:"m²",birimFiyat:684.71,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:398,no:"15.225.1052",tanim:"",birim:"m²",birimFiyat:726.63,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:399,no:"15.225.1053",tanim:"",birim:"m²",birimFiyat:750.83,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:400,no:"15.225.1054",tanim:"",birim:"m²",birimFiyat:792.74,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:401,no:"15.225.1055",tanim:"",birim:"m²",birimFiyat:894.4,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:402,no:"15.225.1056",tanim:"",birim:"m²",birimFiyat:936.31,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:403,no:"15.225.1057",tanim:"",birim:"m²",birimFiyat:995.99,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:404,no:"15.225.1058",tanim:"",birim:"m²",birimFiyat:1097.65,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:405,no:"15.225.1059",tanim:"",birim:"m²",birimFiyat:1163.83,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:406,no:"15.225.1060",tanim:"",birim:"m²",birimFiyat:1212.24,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:407,no:"15.225.1061",tanim:"",birim:"m²",birimFiyat:1313.96,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:408,no:"15.225.1062",tanim:"",birim:"m²",birimFiyat:1415.68,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:409,no:"15.225.1063",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1520.24,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:410,no:"15.225.1064",tanim:"",birim:"m²",birimFiyat:1625.03,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:411,no:"15.225.1065",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1729.85,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:412,no:"15.225.1066",tanim:"",birim:"m²",birimFiyat:1834.78,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:413,no:"15.225.1101",tanim:"",birim:"m²",birimFiyat:731.59,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:414,no:"15.225.1102",tanim:"",birim:"m²",birimFiyat:775.49,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:415,no:"15.225.1103",tanim:"",birim:"m²",birimFiyat:800.59,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:416,no:"15.225.1104",tanim:"",birim:"m²",birimFiyat:844.5,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:417,no:"15.225.1105",tanim:"",birim:"m²",birimFiyat:950.99,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:418,no:"15.225.1106",tanim:"",birim:"m²",birimFiyat:994.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:419,no:"15.225.1107",tanim:"",birim:"m²",birimFiyat:1057.53,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:420,no:"15.225.1108",tanim:"",birim:"m²",birimFiyat:1163.95,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:421,no:"15.225.1109",tanim:"",birim:"m²",birimFiyat:1233.09,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:422,no:"15.225.1110",tanim:"",birim:"m²",birimFiyat:1283.43,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:423,no:"15.225.1111",tanim:"",birim:"m²",birimFiyat:1389.93,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:424,no:"15.225.1112",tanim:"",birim:"m²",birimFiyat:1496.53,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:425,no:"15.225.1113",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1606.24,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:426,no:"15.225.1114",tanim:"",birim:"m²",birimFiyat:1715.93,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:427,no:"15.225.1115",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1825.7,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:428,no:"15.225.1116",tanim:"",birim:"m²",birimFiyat:1935.39,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:429,no:"15.225.1151",tanim:"tutkalı ile)",birim:"m²",birimFiyat:639.65,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:430,no:"15.225.1152",tanim:"tutkalı ile)",birim:"m²",birimFiyat:679.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:431,no:"15.225.1153",tanim:"ile)",birim:"m²",birimFiyat:703.19,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:432,no:"15.225.1154",tanim:"tutkalı ile)",birim:"m²",birimFiyat:743.43,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:433,no:"15.225.1155",tanim:"tutkalı ile)",birim:"m²",birimFiyat:840.78,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:434,no:"15.225.1156",tanim:"tutkalı ile)",birim:"m²",birimFiyat:881.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:435,no:"15.225.1157",tanim:"tutkalı ile)",birim:"m²",birimFiyat:938.11,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:436,no:"15.225.1158",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1035.46,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:437,no:"15.225.1159",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1099.06,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:438,no:"15.225.1160",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1145.66,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:439,no:"15.225.1161",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1243.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:440,no:"15.225.1162",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1340.35,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:441,no:"15.225.1163",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1440.93,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:442,no:"15.225.1164",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1541.48,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:443,no:"15.225.1165",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1642.04,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:444,no:"15.225.1166",tanim:"tutkalı ile)",birim:"m²",birimFiyat:1742.59,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:445,no:"15.225.1301",tanim:"",birim:"m²",birimFiyat:807.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:446,no:"15.225.1302",tanim:"",birim:"m²",birimFiyat:927.58,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:447,no:"15.225.1303",tanim:"",birim:"m²",birimFiyat:1047.19,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:448,no:"15.225.1304",tanim:"",birim:"m²",birimFiyat:1166.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:449,no:"15.225.1305",tanim:"",birim:"m²",birimFiyat:1286.49,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:450,no:"15.225.1306",tanim:"",birim:"m²",birimFiyat:1406.1,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:451,no:"15.225.1307",tanim:"",birim:"m²",birimFiyat:1525.79,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:452,no:"15.225.1401",tanim:"",birim:"m²",birimFiyat:886.94,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:453,no:"15.225.1402",tanim:"",birim:"m²",birimFiyat:984.19,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:454,no:"15.225.1403",tanim:"",birim:"m²",birimFiyat:1037.38,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:455,no:"15.225.1404",tanim:"",birim:"m²",birimFiyat:1134.56,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:456,no:"15.225.1405",tanim:"",birim:"m²",birimFiyat:1373.16,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:457,no:"15.225.1406",tanim:"",birim:"m²",birimFiyat:1470.51,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:458,no:"15.225.1407",tanim:"",birim:"m²",birimFiyat:1620.49,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:459,no:"15.225.1408",tanim:"",birim:"m²",birimFiyat:1859.69,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:460,no:"15.225.1409",tanim:"",birim:"m²",birimFiyat:2009.93,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:461,no:"15.225.1410",tanim:"",birim:"m²",birimFiyat:2116.1,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:462,no:"15.225.1411",tanim:"",birim:"m²",birimFiyat:2356.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:463,no:"15.225.1412",tanim:"",birim:"m²",birimFiyat:2593.44,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:464,no:"15.225.1413",tanim:"",birim:"m²",birimFiyat:2836.54,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:465,no:"15.225.1414",tanim:"",birim:"m²",birimFiyat:3079.66,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:466,no:"15.225.1415",tanim:"",birim:"m²",birimFiyat:3322.86,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:467,no:"15.225.1416",tanim:"",birim:"m²",birimFiyat:3565.98,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:468,no:"15.225.1451",tanim:"",birim:"m²",birimFiyat:969.15,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:469,no:"15.225.1452",tanim:"",birim:"m²",birimFiyat:1071.29,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:470,no:"15.225.1453",tanim:"",birim:"m²",birimFiyat:1126.99,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:471,no:"15.225.1454",tanim:"",birim:"m²",birimFiyat:1229.13,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:472,no:"15.225.1455",tanim:"",birim:"m²",birimFiyat:1480.16,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:473,no:"15.225.1456",tanim:"",birim:"m²",birimFiyat:1582.4,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:474,no:"15.225.1457",tanim:"",birim:"m²",birimFiyat:1740.23,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:475,no:"15.225.1458",tanim:"",birim:"m²",birimFiyat:1991.25,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:476,no:"15.225.1459",tanim:"",birim:"m²",birimFiyat:2149.1,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:477,no:"15.225.1460",tanim:"",birim:"m²",birimFiyat:2260.24,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:478,no:"15.225.1461",tanim:"",birim:"m²",birimFiyat:2511.29,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:479,no:"15.225.1462",tanim:"",birim:"m²",birimFiyat:2762.35,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:480,no:"15.225.1463",tanim:"",birim:"m²",birimFiyat:3017.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:481,no:"15.225.1464",tanim:"",birim:"m²",birimFiyat:3273.36,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:482,no:"15.225.1465",tanim:"",birim:"m²",birimFiyat:3529.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:483,no:"15.225.1466",tanim:"",birim:"m²",birimFiyat:3784.46,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:484,no:"15.225.1611",tanim:"",birim:"m²",birimFiyat:1218.61,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:485,no:"15.225.1612",tanim:"",birim:"m²",birimFiyat:1457.13,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:486,no:"15.225.1613",tanim:"",birim:"m²",birimFiyat:1695.65,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:487,no:"15.225.1614",tanim:"",birim:"m²",birimFiyat:1934.18,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:488,no:"15.225.1615",tanim:"",birim:"m²",birimFiyat:2172.69,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:489,no:"15.225.1616",tanim:"",birim:"m²",birimFiyat:2411.21,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:490,no:"15.225.1617",tanim:"",birim:"m²",birimFiyat:2649.74,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:491,no:"15.225.1618",tanim:"",birim:"m²",birimFiyat:2888.25,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:492,no:"15.225.1711",tanim:"",birim:"m²",birimFiyat:1077.48,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:493,no:"15.225.1712",tanim:"",birim:"m²",birimFiyat:1291.79,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:494,no:"15.225.1713",tanim:"",birim:"m²",birimFiyat:1506.11,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:495,no:"15.225.1714",tanim:"",birim:"m²",birimFiyat:1720.43,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:496,no:"15.225.1715",tanim:"",birim:"m²",birimFiyat:1934.74,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:497,no:"15.225.1811",tanim:"",birim:"m²",birimFiyat:1218.61,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:498,no:"15.225.1812",tanim:"",birim:"m²",birimFiyat:1457.13,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:499,no:"15.225.1813",tanim:"",birim:"m²",birimFiyat:1695.65,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:500,no:"15.225.1814",tanim:"",birim:"m²",birimFiyat:1934.18,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:501,no:"15.225.1815",tanim:"",birim:"m²",birimFiyat:2172.69,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:502,no:"15.225.1901",tanim:"",birim:"m²",birimFiyat:1229.98,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:503,no:"15.225.1902",tanim:"",birim:"m²",birimFiyat:1451.5,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:504,no:"15.225.1903",tanim:"",birim:"m²",birimFiyat:1673.03,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:505,no:"15.225.1904",tanim:"",birim:"m²",birimFiyat:1894.55,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:506,no:"15.225.1905",tanim:"",birim:"m²",birimFiyat:2116.06,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:507,no:"15.225.1906",tanim:"",birim:"m²",birimFiyat:2337.59,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:508,no:"15.225.1907",tanim:"",birim:"m²",birimFiyat:2559.11,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:509,no:"15.225.1908",tanim:"",birim:"m²",birimFiyat:2780.64,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:510,no:"15.225.1909",tanim:"",birim:"m²",birimFiyat:3002.15,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:511,no:"15.225.2001",tanim:"",birim:"m²",birimFiyat:1426.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:512,no:"15.225.2002",tanim:"",birim:"m²",birimFiyat:1686.86,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:513,no:"15.225.2003",tanim:"",birim:"m²",birimFiyat:1946.85,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:514,no:"15.225.2004",tanim:"",birim:"m²",birimFiyat:2206.83,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:515,no:"15.225.2005",tanim:"",birim:"m²",birimFiyat:2466.81,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:516,no:"15.225.2006",tanim:"",birim:"m²",birimFiyat:2730.85,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:517,no:"15.225.2007",tanim:"",birim:"m²",birimFiyat:2986.78,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:518,no:"15.225.2008",tanim:"",birim:"m²",birimFiyat:3246.76,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:519,no:"15.225.2009",tanim:"",birim:"m²",birimFiyat:3506.2,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:520,no:"15.225.2101",tanim:"",birim:"m²",birimFiyat:266.75,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:521,no:"15.225.2102",tanim:"",birim:"m²",birimFiyat:373.13,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:522,no:"15.225.2103",tanim:"",birim:"m²",birimFiyat:431.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:523,no:"15.225.2104",tanim:"",birim:"m²",birimFiyat:506.5,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:524,no:"15.225.2105",tanim:"",birim:"m²",birimFiyat:612.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:525,no:"15.225.2106",tanim:"",birim:"m²",birimFiyat:719.25,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:526,no:"15.225.2107",tanim:"",birim:"m²",birimFiyat:825.63,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:527,no:"15.225.2108",tanim:"",birim:"m²",birimFiyat:932.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:528,no:"15.230.1001",tanim:"",birim:"m²",birimFiyat:525.8,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:529,no:"15.230.1002",tanim:"",birim:"m²",birimFiyat:549.88,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:530,no:"15.230.1003",tanim:"",birim:"m²",birimFiyat:615.84,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:531,no:"15.230.1004",tanim:"",birim:"m²",birimFiyat:647.79,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:532,no:"15.230.1005",tanim:"",birim:"m²",birimFiyat:696.74,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:533,no:"15.230.1006",tanim:"",birim:"m²",birimFiyat:732.55,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:534,no:"15.230.1007",tanim:"",birim:"m²",birimFiyat:856.48,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:535,no:"15.230.1008",tanim:"",birim:"m²",birimFiyat:950.53,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:536,no:"15.230.1101",tanim:"",birim:"m²",birimFiyat:634.85,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:537,no:"15.230.1102",tanim:"",birim:"m²",birimFiyat:741.14,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:538,no:"15.230.1103",tanim:"",birim:"m²",birimFiyat:823.96,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:539,no:"15.230.1201",tanim:"",birim:"m²",birimFiyat:663.94,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:540,no:"15.230.1202",tanim:"",birim:"m²",birimFiyat:723.56,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:541,no:"15.230.1203",tanim:"",birim:"m²",birimFiyat:752.06,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:542,no:"15.230.1204",tanim:"",birim:"m²",birimFiyat:809.06,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:543,no:"15.230.1205",tanim:"",birim:"m²",birimFiyat:888.0,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:544,no:"15.230.1206",tanim:"",birim:"m²",birimFiyat:954.19,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:545,no:"15.230.1207",tanim:"",birim:"m²",birimFiyat:1020.19,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:546,no:"15.230.1208",tanim:"",birim:"m²",birimFiyat:1091.44,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:547,no:"15.230.1301",tanim:"10 cm kalınlığındaki teçhizatlı bimsbeton lento temini ve yerine konulması",birim:"m²",birimFiyat:790.79,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:548,no:"15.230.1302",tanim:"13,5 cm kalınlığındaki teçhizatlı bimsbeton lento temini ve yerine konulması",birim:"m²",birimFiyat:1033.93,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:549,no:"15.230.1303",tanim:"15 cm kalınlığındaki teçhizatlı bimsbeton lento temini ve yerine konulması",birim:"m²",birimFiyat:1141.26,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:550,no:"15.230.1304",tanim:"19 cm kalınlığındaki teçhizatlı bimsbeton lento temini ve yerine konulması",birim:"m²",birimFiyat:1415.44,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:551,no:"15.235.1001",tanim:"",birim:"m²",birimFiyat:911.45,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:552,no:"15.235.1002",tanim:"",birim:"m²",birimFiyat:685.13,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:553,no:"15.235.1003",tanim:"",birim:"m²",birimFiyat:735.13,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:554,no:"15.235.1004",tanim:"",birim:"m²",birimFiyat:1042.33,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:555,no:"15.235.1005",tanim:"",birim:"m²",birimFiyat:754.7,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:556,no:"15.235.1024",tanim:"",birim:"m²",birimFiyat:623.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:557,no:"15.235.1025",tanim:"",birim:"m²",birimFiyat:681.23,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:558,no:"15.235.1027",tanim:"",birim:"m²",birimFiyat:759.8,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:559,no:"15.235.1028",tanim:"",birim:"m²",birimFiyat:825.75,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:560,no:"15.235.1031",tanim:"",birim:"m²",birimFiyat:906.59,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:561,no:"15.235.1032",tanim:"",birim:"m²",birimFiyat:974.54,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:562,no:"15.235.1033",tanim:"",birim:"m²",birimFiyat:1042.49,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:563,no:"15.235.1034",tanim:"",birim:"m²",birimFiyat:1111.68,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:564,no:"15.235.1035",tanim:"",birim:"m²",birimFiyat:1182.84,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:565,no:"15.235.1036",tanim:"",birim:"m²",birimFiyat:1254.03,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:566,no:"15.235.1037",tanim:"",birim:"m²",birimFiyat:1325.19,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:567,no:"15.235.1038",tanim:"",birim:"m²",birimFiyat:1394.38,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:568,no:"15.235.1039",tanim:"",birim:"m²",birimFiyat:1465.54,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:569,no:"15.235.1043",tanim:"",birim:"m²",birimFiyat:1748.24,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:570,no:"15.235.1047",tanim:"",birim:"m²",birimFiyat:2024.41,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:571,no:"15.235.1051",tanim:"15 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:676.59,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:572,no:"15.235.1052",tanim:"17,5 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:765.68,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:573,no:"15.235.1053",tanim:"20 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:870.79,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:574,no:"15.235.1054",tanim:"22,5 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:967.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:575,no:"15.235.1055",tanim:"25 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:1066.99,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:576,no:"15.235.1056",tanim:"27,5 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:1164.09,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:577,no:"15.235.1057",tanim:"30 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:1261.19,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:578,no:"15.235.1058",tanim:"32,5 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:1358.28,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:579,no:"15.235.1059",tanim:"35 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:1457.38,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:580,no:"15.235.1060",tanim:"37,5 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:1554.48,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:581,no:"15.235.1061",tanim:"40 cm yüksekliğindeki EPS katkılı bloklar ile asmolen döşeme yapılması",birim:"m²",birimFiyat:1651.58,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:582,no:"15.240.1001",tanim:"",birim:"m²",birimFiyat:505.29,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:583,no:"15.240.1002",tanim:"",birim:"m²",birimFiyat:597.63,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:584,no:"15.240.1003",tanim:"",birim:"m²",birimFiyat:661.83,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:585,no:"15.245.1001",tanim:"150 gr/m² ağırlıkta geotekstil keçe serilmesi",birim:"m²",birimFiyat:52.75,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:586,no:"15.245.1002",tanim:"250 gr/m² ağırlıkta geotekstil keçe serilmesi",birim:"m²",birimFiyat:58.25,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:587,no:"15.245.1003",tanim:"500 gr/m² ağırlıkta geotekstil keçe serilmesi",birim:"m²",birimFiyat:73.38,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:588,no:"15.250.1011",tanim:"200 kg çimento dozlu tesviye tabakası yapılması",birim:"m²",birimFiyat:335.74,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:589,no:"15.250.1111",tanim:"2.5 cm kalınlığında 400 kg çimento dozlu şap yapılması",birim:"m²",birimFiyat:449.53,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:590,no:"15.250.1112",tanim:"2.5 cm kalınlığında 450 kg çimento dozlu şap yapılması",birim:"m²",birimFiyat:450.69,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:591,no:"15.250.1113",tanim:"2.5 cm kalınlığında 500 kg çimento dozlu şap yapılması",birim:"m²",birimFiyat:463.95,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:592,no:"15.250.1114",tanim:"Makina ile ortalama 2,5 cm kalınlıkta alçı esaslı şap yapılması",birim:"m²",birimFiyat:385.59,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:593,no:"15.255.1001",tanim:"kalınlıkta plastomer esaslı (-5 °C soğukta bükülmeli) polyester keçe taşıyıcılı polimer",birim:"m²",birimFiyat:586.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:594,no:"15.255.1002",tanim:"kalınlıkta plastomer esaslı (-10 °C soğukta bükülmeli) polyester keçe taşıyıcılı polimer",birim:"m²",birimFiyat:611.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:595,no:"15.255.1003",tanim:"kalınlıkta elastomer esaslı (-20 °C soğukta bükülmeli) polyester keçe taşıyıcılı polimer",birim:"m²",birimFiyat:679.45,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:596,no:"15.255.1004",tanim:"",birim:"m²",birimFiyat:611.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:597,no:"15.255.1005",tanim:"",birim:"m²",birimFiyat:640.64,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:598,no:"15.255.1006",tanim:"",birim:"m²",birimFiyat:726.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:599,no:"15.255.1007",tanim:"",birim:"m²",birimFiyat:655.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:600,no:"15.255.1008",tanim:"",birim:"m²",birimFiyat:683.76,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:601,no:"15.255.1009",tanim:"",birim:"m²",birimFiyat:767.14,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:602,no:"15.255.1010",tanim:"kalınlıkta plastomer esaslı (-5 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü",birim:"m²",birimFiyat:629.14,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:603,no:"15.255.1011",tanim:"kalınlıkta plastomer esaslı (-10 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü",birim:"m²",birimFiyat:655.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:604,no:"15.255.1012",tanim:"kalınlıkta elastomer esaslı (-20 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü",birim:"m²",birimFiyat:715.39,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:605,no:"15.255.1013",tanim:"mm kalınlıkta plastomer esaslı (-5 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü",birim:"m²",birimFiyat:655.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:606,no:"15.255.1014",tanim:"mm kalınlıkta plastomer esaslı (-10 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü",birim:"m²",birimFiyat:683.76,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:607,no:"15.255.1015",tanim:"mm kalınlıkta elastomer esaslı (-20 °C soğukta bükülmeli) polyester keçe taşıyıcılı bir yüzü",birim:"m²",birimFiyat:762.83,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:608,no:"15.255.1016",tanim:"",birim:"m²",birimFiyat:408.7,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:609,no:"15.255.1017",tanim:"",birim:"m²",birimFiyat:424.51,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:610,no:"15.255.1018",tanim:"",birim:"m²",birimFiyat:467.64,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:611,no:"15.255.1019",tanim:"",birim:"m²",birimFiyat:424.51,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:612,no:"15.255.1020",tanim:"",birim:"m²",birimFiyat:502.14,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:613,no:"15.255.1021",tanim:"",birim:"m²",birimFiyat:310.95,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:614,no:"15.255.1022",tanim:"",birim:"m²",birimFiyat:336.83,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:615,no:"15.255.1023",tanim:"",birim:"m²",birimFiyat:322.45,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:616,no:"15.255.1024",tanim:"",birim:"m²",birimFiyat:351.2,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:617,no:"15.255.1025",tanim:"",birim:"m²",birimFiyat:346.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:618,no:"15.255.1026",tanim:"",birim:"m²",birimFiyat:394.33,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:619,no:"15.260.1001",tanim:"",birim:"m²",birimFiyat:430.4,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:620,no:"15.260.1002",tanim:"",birim:"m²",birimFiyat:509.15,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:621,no:"15.260.1003",tanim:"1,5 mm kalınlıkta, PVC esaslı, (UV dayanımlı, donatılı) jeomembran ile su yalıtımı yapılması",birim:"m²",birimFiyat:456.65,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:622,no:"15.260.1004",tanim:"2 mm kalınlıkta, PVC esaslı, (UV dayanımlı, donatılı) jeomembran ile su yalıtımı yapılması",birim:"m²",birimFiyat:535.4,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:623,no:"15.260.1005",tanim:"",birim:"m²",birimFiyat:391.03,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:624,no:"15.260.1006",tanim:"",birim:"m²",birimFiyat:450.09,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:625,no:"15.260.1007",tanim:"",birim:"m²",birimFiyat:410.71,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:626,no:"15.260.1008",tanim:"2 mm kalınlıkta, HDPE esaslı, (UV dayanımlı, donatılı) jeomembran ile su yalıtımı yapılması",birim:"m²",birimFiyat:482.9,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:627,no:"15.260.1009",tanim:"",birim:"m²",birimFiyat:391.03,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:628,no:"15.260.1010",tanim:"",birim:"m²",birimFiyat:450.09,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:629,no:"15.260.1011",tanim:"1,5 mm kalınlıkta, Termoset EPDM esaslı, jeomembran ile su yalıtımı yapılması",birim:"m²",birimFiyat:692.9,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:630,no:"15.260.1012",tanim:"2 mm kalınlıkta, Termoset EPDM esaslı, jeomembran ile su yalıtımı yapılması",birim:"m²",birimFiyat:876.65,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:631,no:"15.260.1013",tanim:"1,5 mm kalınlıkta, TPO esaslı, (UV dayanımlı, donatılı) jeomembran ile su yalıtımı yapılması",birim:"m²",birimFiyat:541.96,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:632,no:"15.260.1014",tanim:"2 mm kalınlıkta, TPO esaslı, (UV dayanımlı, donatılı) jeomembran ile su yalıtımı yapılması",birim:"m²",birimFiyat:653.53,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:633,no:"15.265.1001",tanim:"3 mm kalınlıkta, HDPE levhalarla, su yalıtımı yapılması",birim:"m²",birimFiyat:603.21,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:634,no:"15.265.1002",tanim:"4 mm kalınlıkta, HDPE levhalarla, su yalıtımı yapılması",birim:"m²",birimFiyat:736.03,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:635,no:"15.265.1003",tanim:"5 mm kalınlıkta, HDPE levhalarla, su yalıtımı yapılması",birim:"m²",birimFiyat:862.28,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:636,no:"15.265.1004",tanim:"3 mm kalınlıkta, PP levhalarla, su yalıtımı yapılması",birim:"m²",birimFiyat:577.03,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:637,no:"15.265.1005",tanim:"4 mm kalınlıkta, PP levhalarla, su yalıtımı yapılması",birim:"m²",birimFiyat:709.85,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:638,no:"15.265.1006",tanim:"5 mm kalınlıkta, PP levhalarla, su yalıtımı yapılması",birim:"m²",birimFiyat:849.24,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:639,no:"15.270.1001",tanim:"",birim:"m²",birimFiyat:586.29,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:640,no:"15.270.1002",tanim:"",birim:"m²",birimFiyat:619.41,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:641,no:"15.270.1003",tanim:"",birim:"m²",birimFiyat:750.35,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:642,no:"15.270.1004",tanim:"",birim:"m²",birimFiyat:783.48,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:643,no:"15.270.1005",tanim:"",birim:"m²",birimFiyat:530.04,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:644,no:"15.270.1006",tanim:"",birim:"m²",birimFiyat:563.16,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:645,no:"15.270.1007",tanim:"",birim:"m²",birimFiyat:637.85,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:646,no:"15.270.1008",tanim:"",birim:"m²",birimFiyat:670.98,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:647,no:"15.270.1009",tanim:"",birim:"m²",birimFiyat:473.85,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:648,no:"15.270.1010",tanim:"",birim:"m²",birimFiyat:506.98,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:649,no:"15.270.1011",tanim:"",birim:"m²",birimFiyat:562.95,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:650,no:"15.270.1012",tanim:"",birim:"m²",birimFiyat:596.08,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:651,no:"15.270.1101",tanim:"",birim:"m²",birimFiyat:1469.4,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:652,no:"15.270.1111",tanim:"",birim:"m²",birimFiyat:2170.65,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:653,no:"15.270.1202",tanim:"Örgülü Geotekstil, Üst Tabaka 200 gr/m² PP Örgüsüz Geotekstil, Toplam Ağırlık 5500",birim:"m²",birimFiyat:293.75,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:654,no:"15.270.1203",tanim:"Örgülü Geotekstil, Üst Tabaka 200 gr/m² PP Örgüsüz Geotekstil, Toplam Ağırlık 6500",birim:"m²",birimFiyat:300.63,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:655,no:"15.275.1011",tanim:"Taş duvar yüzeylerine gömme oluklu derz yapılması",birim:"m²",birimFiyat:265.98,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:656,no:"15.275.1012",tanim:"Taş duvar yüzeylerine kabartma derz yapılması",birim:"m²",birimFiyat:290.35,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:657,no:"15.275.1111",tanim:"250/350 kg çimento dozlu kaba ve ince harçla sıva yapılması (dış cephe sıvası)",birim:"m²",birimFiyat:744.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:658,no:"15.275.1112",tanim:"200/250 kg kireç/çimento karışımı kaba ve ince harçla sıva yapılması (iç cephe sıvası)",birim:"m²",birimFiyat:661.76,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:659,no:"15.275.1113",tanim:"250/350 kg kireç/çimento karışımı kaba ve ince harçla sıva yapılması (tavan sıvası)",birim:"m²",birimFiyat:693.65,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:660,no:"15.275.1114",tanim:"250/350 kg çimento dozlu kaba ve ince harçla serpme (çarpma) sıva yapılması",birim:"m²",birimFiyat:517.38,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:661,no:"15.275.1115",tanim:"350 kg çimento dozlu harçla tek kat ince sıva yapılması",birim:"m²",birimFiyat:483.61,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:662,no:"15.275.1116",tanim:"250 kg çimento dozlu harç ile kaba sıva yapılması",birim:"m²",birimFiyat:443.51,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:663,no:"15.275.1117",tanim:"200 kg Çimento ve Kireç karşımı harç ile kaba sıva yapılması (iç cephe)",birim:"m²",birimFiyat:454.16,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:664,no:"15.275.9992",tanim:"Metal kapı kasa arkalarının beton harcı ile doldurulması",birim:"m²",birimFiyat:623.56,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:665,no:"15.280.1009",tanim:"Perlitli sıva alçısı ve saten alçı ile kaplama yapılması (Beton, tuğla duvar vb. yüzeylere)",birim:"m²",birimFiyat:694.59,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:666,no:"15.280.1010",tanim:"",birim:"m²",birimFiyat:245.86,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:667,no:"15.280.1011",tanim:"Saten alçı kaplaması yapılması (ortalama 1 mm kalınlık)",birim:"m²",birimFiyat:139.71,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:668,no:"15.280.1012",tanim:"Makina alçısı ile tavanlara 15 mm kalınlığında tek kat alçı sıva yapılması",birim:"m²",birimFiyat:468.55,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:669,no:"15.280.1013",tanim:"",birim:"m²",birimFiyat:505.15,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:670,no:"15.285.1001",tanim:"",birim:"m²",birimFiyat:653.86,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:671,no:"15.285.1002",tanim:"",birim:"m²",birimFiyat:877.66,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:672,no:"15.285.1003",tanim:"",birim:"m²",birimFiyat:1101.48,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:673,no:"15.285.1011",tanim:"",birim:"m²",birimFiyat:667.61,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:674,no:"15.285.1012",tanim:"",birim:"m²",birimFiyat:898.29,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:675,no:"15.285.1013",tanim:"",birim:"m²",birimFiyat:1128.98,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:676,no:"15.300.1001",tanim:"Ahşaptan oturtma çatı yapılması (çatı örtüsü altı tahta kaplamalı)",birim:"m²",birimFiyat:1417.89,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:677,no:"15.300.1002",tanim:"Ahşaptan oturtma çatı yapılması (çatı örtüsünün altı OSB/3 kaplamalı)",birim:"m²",birimFiyat:1566.83,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:678,no:"15.300.1003",tanim:"Ahşaptan makaslı çatı yapılması",birim:"m³",birimFiyat:28881.34,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:679,no:"15.300.1004",tanim:"Rendeli ahşaptan makaslı çatı yapılması",birim:"m³",birimFiyat:29746.83,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:680,no:"15.300.1005",tanim:"Çatı üzerine tahta kaplama yapılması",birim:"m²",birimFiyat:611.06,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:681,no:"15.300.1006",tanim:"Çatı üzerine OSB/3 kaplama yapılması",birim:"m²",birimFiyat:684.69,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:682,no:"15.300.1007",tanim:"Rendeli ahşaptan saçak altı ve alın kaplaması yapılması",birim:"m²",birimFiyat:1031.53,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:683,no:"15.305.1001",tanim:"",birim:"m²",birimFiyat:1577.38,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:684,no:"15.305.1002",tanim:"",birim:"m²",birimFiyat:1519.63,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:685,no:"15.305.1003",tanim:"",birim:"m²",birimFiyat:1070.5,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:686,no:"15.305.1004",tanim:"",birim:"m²",birimFiyat:1029.81,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:687,no:"15.305.1201",tanim:"Renksiz beton kiremitler ile çatı örtüsü yapılması (çift latalı sistem)",birim:"m²",birimFiyat:949.75,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:688,no:"15.305.1202",tanim:"Demir oksit boyalı beton kiremitler ile çatı örtüsü yapılması (çift latalı sistem)",birim:"m²",birimFiyat:994.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:689,no:"15.305.1203",tanim:"",birim:"m²",birimFiyat:1041.63,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:690,no:"15.305.1204",tanim:"Renksiz beton mahya kiremitleri ile mahya yapılması",birim:"m",birimFiyat:567.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:691,no:"15.305.1205",tanim:"Demir oksit boyalı beton mahya kiremitleri ile mahya yapılması",birim:"m",birimFiyat:604.13,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:692,no:"15.305.1207",tanim:"Renksiz perlitli beton kiremitler ile çatı örtüsü yapılması (çift latalı sistem)",birim:"m²",birimFiyat:910.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:693,no:"15.305.1208",tanim:"Demir oksit boyalı perlitli beton kiremitler ile çatı örtüsü yapılması (çift latalı sistem)",birim:"m²",birimFiyat:949.75,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:694,no:"15.305.1209",tanim:"",birim:"m²",birimFiyat:994.38,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:695,no:"15.305.1210",tanim:"Renksiz perlitli beton mahya kiremitleri ile mahya yapılması",birim:"m",birimFiyat:552.94,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:696,no:"15.305.1213",tanim:"yapışkanlı, UV dayanımlı duvar baca dibi bandı ile duvar, baca dibi vb. yerlerde su yalıtımı",birim:"m",birimFiyat:472.31,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:697,no:"15.305.1214",tanim:"Alüminyum baskı çıtası ve poliüretan mastik ile yalıtım bitişlerinde sızdırmazlık sağlanması",birim:"m",birimFiyat:294.83,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:698,no:"15.305.1215",tanim:"",birim:"m",birimFiyat:459.0,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:699,no:"15.310.1001",tanim:"12 no'lu çinko levhadan 150 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.",birim:"m",birimFiyat:1085.9,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:700,no:"15.310.1002",tanim:"12 no'lu çinko levhadan 120 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.",birim:"m",birimFiyat:973.39,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:701,no:"15.310.1003",tanim:"12 no'lu çinko levhadan 100 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.",birim:"m",birimFiyat:898.98,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:702,no:"15.310.1004",tanim:"10 no'lu çinko levhadan 100 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.",birim:"m",birimFiyat:820.96,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:703,no:"15.310.1005",tanim:"10 no'lu çinko levhadan 80 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.",birim:"m",birimFiyat:776.16,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:704,no:"15.310.1006",tanim:"12 no'lu çinko levhadan 80 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.",birim:"m",birimFiyat:842.91,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:705,no:"15.310.1007",tanim:"10 no'lu çinko levhadan 75 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.",birim:"m",birimFiyat:747.55,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:706,no:"15.310.1008",tanim:"10 no'lu çinko levhadan 70 mm çapında düşey yağmur borusu yapılması ve yerine tespiti.",birim:"m",birimFiyat:707.86,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:707,no:"15.310.1101",tanim:"14 no'lu çinko levhadan 240 mm çapında yağmur oluğu yapılması ve yerine tespiti",birim:"m",birimFiyat:2264.46,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:708,no:"15.310.1102",tanim:"12 no'lu çinko levhadan 185 mm çapında yağmur oluğu yapılması ve yerine tespiti.",birim:"m",birimFiyat:1877.18,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:709,no:"15.310.1103",tanim:"12 no'lu çinko levhadan 155 mm çapında yağmur oluğu yapılması ve yerine tespiti.",birim:"m",birimFiyat:1748.6,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:710,no:"15.310.1104",tanim:"12 no'lu çinko levhadan 130 mm çapında yağmur oluğu yapılması ve yerine tespiti.",birim:"m",birimFiyat:1608.45,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:711,no:"15.310.1105",tanim:"12 no'lu çinko levhadan 110 mm çapında yağmur oluğu yapılması ve yerine tespiti.",birim:"m",birimFiyat:1534.06,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:712,no:"15.310.1106",tanim:"12 no'lu çinko levhadan 90 mm çapında yağmur oluğu yapılması ve yerine tespiti.",birim:"m",birimFiyat:1438.16,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:713,no:"15.310.1201",tanim:"14 no.lu çinko levhadan eğimli çatı deresi yapılması ve yerine konulması",birim:"m",birimFiyat:1344.89,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:714,no:"15.310.1202",tanim:"14 no.lu çinko levhadan oluk halinde yatay çatı deresi yapılması ve yerine konulması",birim:"m",birimFiyat:2626.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:715,no:"15.310.1203",tanim:"",birim:"adet",birimFiyat:3012.39,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:716,no:"15.310.1204",tanim:"14 no.lu çinko levhadan atika duvar arkası için çatı deresi yapılması ve yerine konulması",birim:"m",birimFiyat:2818.48,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:717,no:"15.310.1205",tanim:"",birim:"m",birimFiyat:1065.46,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:718,no:"15.310.1206",tanim:"",birim:"m²",birimFiyat:1919.03,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:719,no:"15.310.1207",tanim:"12 no.lu çinko levhadan pencere denizliği yapılması ve yerine konulması",birim:"m",birimFiyat:981.49,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:720,no:"15.310.1208",tanim:"12 no.lu çinko levhadan baca temizleme kutusu yapılması ve yerine konulması",birim:"adet",birimFiyat:575.98,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:721,no:"15.310.1209",tanim:"12 no.lu çinko levhadan soba deliği ağızlığı ve kapağı yapılması ve yerine konulması",birim:"adet",birimFiyat:412.8,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:722,no:"15.310.1301",tanim:"0.50 mm bakır levhadan 125 mm çapında düşey yağmur borusu yapılması ve yerine tespiti",birim:"m",birimFiyat:1955.96,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:723,no:"15.310.1302",tanim:"",birim:"m",birimFiyat:3108.18,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:724,no:"15.310.1303",tanim:"0,50 mm bakır levhadan çatı deresi yapılması ve yerine tesbiti",birim:"m",birimFiyat:3155.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:725,no:"15.310.1304",tanim:"0,50 mm bakır levhadan oluk halinde çatı deresi yapılması ve yerine tespiti",birim:"m",birimFiyat:5487.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:726,no:"15.310.1305",tanim:"",birim:"adet",birimFiyat:4874.18,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:727,no:"15.310.1306",tanim:"0,50 mm bakır levhadan atika duvarı arkasına çatı deresi yapılması ve yerine konulması",birim:"m",birimFiyat:5778.75,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:728,no:"15.310.1307",tanim:"",birim:"m",birimFiyat:2253.94,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:729,no:"15.310.1308",tanim:"",birim:"m²",birimFiyat:4563.41,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:730,no:"15.310.1309",tanim:"0,50 mm bakır levhadan pencere denizliği yapılması ve yerine konulması",birim:"m",birimFiyat:1980.81,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:731,no:"15.315.1001",tanim:"Ø 70 mm çapında bir ucu muflu sert PVC yağmur borusu temini ve yerine tesbiti",birim:"m",birimFiyat:207.41,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:732,no:"15.315.1002",tanim:"Ø 100 mm çapında bir ucu muflu sert PVC yağmur borusu temini ve yerine tesbiti",birim:"m",birimFiyat:268.14,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:733,no:"15.315.1003",tanim:"Ø 125 mm çapında bir ucu muflu sert PVC yağmur borusu temini ve yerine tesbiti",birim:"m",birimFiyat:310.75,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:734,no:"15.315.1004",tanim:"Ø 100 mm çapında sert PVC yağmur oluğu temini ve yerine tesbiti",birim:"m",birimFiyat:371.1,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:735,no:"15.315.1005",tanim:"Ø 150 mm çapında sert PVC yağmur oluğu temini ve yerine tesbiti",birim:"m",birimFiyat:437.71,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:736,no:"15.320.1001",tanim:"mm kalınlıkta boyalı galvanizli sac ve altı 0.40 mm kalınlıkta boyalı galvanizli sac) çatı",birim:"m²",birimFiyat:1459.66,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:737,no:"15.320.1002",tanim:"mm kalınlıkta PVC membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli ile",birim:"m²",birimFiyat:2015.91,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:738,no:"15.320.1003",tanim:"mm kalınlıkta TPO membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli ile",birim:"m²",birimFiyat:2075.91,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:739,no:"15.320.1004",tanim:"0.50 mm kalınlıkta boyalı galvanizli sac ve altı 0.40 mm kalınlıkta boyalı galvanizli sac) çatı",birim:"m²",birimFiyat:1572.16,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:740,no:"15.320.1005",tanim:"1.20 mm kalınlıkta PVC membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli",birim:"m²",birimFiyat:2165.91,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:741,no:"15.320.1006",tanim:"1.20 mm kalınlıkta TPO membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli",birim:"m²",birimFiyat:2180.91,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:742,no:"15.320.1007",tanim:"mm kalınlıkta, altı 0.50 mm kalınlıkta naturel, gofrajlı alüminyum) çatı paneli ile çatı örtüsü",birim:"m²",birimFiyat:1684.66,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:743,no:"15.320.1008",tanim:"mm kalınlıkta boyalı galvanizli sac, altı 0.40 mm kalınlıkta boyalı galvanizli sac) çatı paneli",birim:"m²",birimFiyat:1437.16,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:744,no:"15.320.1009",tanim:"mm kalınlıkta boyalı galvanizli sac, altı 0.40 mm kalınlıkta, naturel, gofrajlı alüminyum) çatı",birim:"m²",birimFiyat:1557.16,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:745,no:"15.320.1010",tanim:"galvanizli sac ve altı 0.50 mm kalınlıkta boyalı galvanizli sac) çatı paneli ile çatı örtüsü",birim:"m²",birimFiyat:1854.98,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:746,no:"15.320.1011",tanim:"",birim:"m²",birimFiyat:2415.6,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:747,no:"15.320.1012",tanim:"",birim:"m²",birimFiyat:2498.1,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:748,no:"15.320.1013",tanim:"",birim:"m²",birimFiyat:2543.1,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:749,no:"15.320.1014",tanim:"kalınlıkta PVC membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli ile çatı",birim:"m²",birimFiyat:2640.6,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:750,no:"15.320.1015",tanim:"kalınlıkta TPO membranlı, altı 0.60 mm kalınlıkta boyalı galvanizli sac) çatı paneli ile çatı",birim:"m²",birimFiyat:2693.1,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:751,no:"15.325.1001",tanim:"Ahşap çatı üzerine 0,50 mm kalınlıkta 10 nolu çinkodan çatı örtüsü yapılması",birim:"m²",birimFiyat:2398.81,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:752,no:"15.325.1002",tanim:"Ahşap çatı üzerine 0,50 mm kalınlıkta bakır levhadan çatı örtüsü yapılması",birim:"m²",birimFiyat:5560.13,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:753,no:"15.325.1003",tanim:"Ahşap çatı üzerine 0,66 mm bakır levhadan çatı örtüsü yapılması",birim:"m²",birimFiyat:6871.31,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:754,no:"15.325.1004",tanim:"",birim:"m²",birimFiyat:2329.66,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:755,no:"15.325.1005",tanim:"",birim:"m²",birimFiyat:1190.03,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:756,no:"15.325.1006",tanim:"",birim:"m²",birimFiyat:1160.5,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:757,no:"15.325.1007",tanim:"üzerine 0.70 mm kalınlığında trapezoidal alüminyum levhalar (EN AW 3003 Al-Mn1 Cu) ile",birim:"m²",birimFiyat:1364.4,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:758,no:"15.325.1008",tanim:"",birim:"m²",birimFiyat:1470.75,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:759,no:"15.325.1009",tanim:"",birim:"m²",birimFiyat:654.61,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:760,no:"15.325.1101",tanim:"",birim:"m²",birimFiyat:544.44,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:761,no:"15.325.1102",tanim:"",birim:"m²",birimFiyat:464.38,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:762,no:"15.325.1103",tanim:"mevcut çatı döşemesi üzerine 0.50 mm kalınlıkta sıcak daldırma galvanizli, oluklu/trapez sac",birim:"m²",birimFiyat:697.11,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:763,no:"15.325.1104",tanim:"mevcut çatı döşemesi üzerine elyaflı çimentodan yapılmış oluklu levhalarla çatı örtüsü",birim:"m²",birimFiyat:631.94,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:764,no:"15.325.1105",tanim:"mevcut çatı döşemesi üzerine her renkte oluklu bitümlü levhalarla çatı örtüsü yapılması",birim:"m²",birimFiyat:554.99,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:765,no:"15.325.1106",tanim:"",birim:"m²",birimFiyat:539.86,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:766,no:"15.325.1107",tanim:"Betonarme çatı üzerine kurşun levha ile çatı örtüsü yapılması.",birim:"kg",birimFiyat:212.64,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:767,no:"15.325.1108",tanim:"",birim:"m²",birimFiyat:846.5,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:768,no:"15.325.1109",tanim:"",birim:"m²",birimFiyat:658.36,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:769,no:"15.325.1110",tanim:"",birim:"m²",birimFiyat:593.19,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:770,no:"15.330.1001",tanim:"",birim:"m²",birimFiyat:232.56,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:771,no:"15.330.1002",tanim:"",birim:"m²",birimFiyat:246.31,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:772,no:"15.330.1003",tanim:"",birim:"m²",birimFiyat:239.44,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:773,no:"15.330.1004",tanim:"",birim:"m²",birimFiyat:379.2,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:774,no:"15.330.1005",tanim:"",birim:"m²",birimFiyat:406.7,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:775,no:"15.330.1006",tanim:"",birim:"m²",birimFiyat:402.58,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:776,no:"15.330.1007",tanim:"",birim:"m²",birimFiyat:447.95,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:777,no:"15.330.1008",tanim:"",birim:"m²",birimFiyat:368.2,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:778,no:"15.330.1009",tanim:"",birim:"m²",birimFiyat:392.95,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:779,no:"15.330.1010",tanim:"",birim:"m²",birimFiyat:378.75,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:780,no:"15.330.1011",tanim:"",birim:"m²",birimFiyat:476.86,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:781,no:"15.340.9951",tanim:"",birim:"m²",birimFiyat:87.75,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:782,no:"15.340.9952",tanim:"",birim:"m²",birimFiyat:154.69,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:783,no:"15.340.9953",tanim:"",birim:"m²",birimFiyat:226.88,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:784,no:"15.340.9961",tanim:"",birim:"m²",birimFiyat:135.0,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:785,no:"15.340.9962",tanim:"",birim:"m²",birimFiyat:167.81,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:786,no:"15.340.9963",tanim:"",birim:"m²",birimFiyat:226.88,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:787,no:"15.341.1001",tanim:"levhalar (EPS ) ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası yapılması",birim:"m²",birimFiyat:1000.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:788,no:"15.341.1002",tanim:"levhalar (EPS) ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası yapılması",birim:"m²",birimFiyat:1022.46,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:789,no:"15.341.1003",tanim:"levhalar (EPS) ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası yapılması",birim:"m²",birimFiyat:1065.78,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:790,no:"15.341.1004",tanim:"levhalar ile (EPS) dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası yapılması",birim:"m²",birimFiyat:1109.09,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:791,no:"15.341.1021",tanim:"ekspande polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası",birim:"m²",birimFiyat:1020.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:792,no:"15.341.1022",tanim:"ekspande polistren (EPS) levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım",birim:"m²",birimFiyat:1046.09,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:793,no:"15.341.1023",tanim:"ekspande polistren levhalar (EPS) ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım",birim:"m²",birimFiyat:1097.28,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:794,no:"15.341.1024",tanim:"ekspande polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı yalıtım sıvası",birim:"m²",birimFiyat:1148.46,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:795,no:"15.341.1041",tanim:"",birim:"m²",birimFiyat:144.85,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:796,no:"15.341.1042",tanim:"",birim:"m²",birimFiyat:157.98,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:797,no:"15.341.1043",tanim:"",birim:"m²",birimFiyat:171.1,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:798,no:"15.341.1044",tanim:"",birim:"m²",birimFiyat:187.5,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:799,no:"15.341.1061",tanim:"",birim:"m²",birimFiyat:190.79,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:800,no:"15.341.1062",tanim:"",birim:"m²",birimFiyat:230.16,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:801,no:"15.341.1063",tanim:"",birim:"m²",birimFiyat:236.73,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:802,no:"15.341.1064",tanim:"",birim:"m²",birimFiyat:282.66,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:803,no:"15.341.1081",tanim:"",birim:"m²",birimFiyat:986.11,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:804,no:"15.341.1082",tanim:"",birim:"m²",birimFiyat:1022.21,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:805,no:"15.341.1083",tanim:"Polistren levhalar (EPS) ile duvarlarda içten ısı yalıtımı ve üzerine ısı yalıtım sıvası",birim:"m²",birimFiyat:1015.65,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:806,no:"15.341.1084",tanim:"Ekspande Polistren levhalar (EPS) ile duvarlarda içten ısı yalıtımı ve üzerine ısı yalıtım",birim:"m²",birimFiyat:1068.15,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:807,no:"15.341.1101",tanim:"",birim:"m²",birimFiyat:292.5,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:808,no:"15.341.1102",tanim:"",birim:"m²",birimFiyat:338.44,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:809,no:"15.341.1103",tanim:"",birim:"m²",birimFiyat:364.69,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:810,no:"15.341.1104",tanim:"",birim:"m²",birimFiyat:430.31,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:811,no:"15.341.1121",tanim:"",birim:"m²",birimFiyat:417.19,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:812,no:"15.341.1122",tanim:"",birim:"m²",birimFiyat:489.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:813,no:"15.341.1123",tanim:"",birim:"m²",birimFiyat:522.19,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:814,no:"15.341.1124",tanim:"",birim:"m²",birimFiyat:620.63,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:815,no:"15.341.1141",tanim:"",birim:"m²",birimFiyat:342.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:816,no:"15.341.1142",tanim:"",birim:"m²",birimFiyat:426.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:817,no:"15.341.2001",tanim:"pürüzlü kanallı ekstrüde polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine",birim:"m²",birimFiyat:1207.53,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:818,no:"15.341.2002",tanim:"pürüzlü kanallı ekstrüde polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı",birim:"m²",birimFiyat:1345.34,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:819,no:"15.341.2003",tanim:"pürüzlü kanallı ekstrüde polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı",birim:"m²",birimFiyat:1312.53,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:820,no:"15.341.2004",tanim:"pürüzlü kanallı ekstrüde polistren levhalar ile dış duvarlarda dıştan ısı yalıtımı ve üzerine ısı",birim:"m²",birimFiyat:1496.28,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:821,no:"15.341.2021",tanim:"",birim:"m²",birimFiyat:269.54,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:822,no:"15.341.2022",tanim:"",birim:"m²",birimFiyat:302.35,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:823,no:"15.341.2023",tanim:"",birim:"m²",birimFiyat:314.81,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:824,no:"15.341.2024",tanim:"",birim:"m²",birimFiyat:354.19,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:825,no:"15.341.2025",tanim:"",birim:"m²",birimFiyat:377.81,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:826,no:"15.341.2026",tanim:"",birim:"m²",birimFiyat:489.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:827,no:"15.341.2041",tanim:"kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras çatı",birim:"m²",birimFiyat:328.6,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:828,no:"15.341.2042",tanim:"cm kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras",birim:"m²",birimFiyat:385.69,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:829,no:"15.341.2043",tanim:"cm kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras",birim:"m²",birimFiyat:499.88,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:830,no:"15.341.2044",tanim:"cm kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras",birim:"m²",birimFiyat:614.06,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:831,no:"15.341.2045",tanim:"kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras çatı",birim:"m²",birimFiyat:322.04,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:832,no:"15.341.2046",tanim:"kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras çatı",birim:"m²",birimFiyat:377.81,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:833,no:"15.341.2047",tanim:"kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras çatı",birim:"m²",birimFiyat:489.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:834,no:"15.341.2048",tanim:"cm kalınlıkta (XPS levhalar yüklenebilen) levhalar ile yatayda (geleneksel gezilebilir teras",birim:"m²",birimFiyat:600.94,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:835,no:"15.341.2061",tanim:"cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilmeyen ters teras çatılarda) ısı",birim:"m²",birimFiyat:600.94,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:836,no:"15.341.2062",tanim:"cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilmeyen ters teras çatılarda) ısı",birim:"m²",birimFiyat:594.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:837,no:"15.341.2063",tanim:"cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilmeyen ters teras çatılarda) ısı",birim:"m²",birimFiyat:555.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:838,no:"15.341.2064",tanim:"cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilen ters teras çatılarda) ısı yalıtımı",birim:"m²",birimFiyat:646.88,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:839,no:"15.341.2065",tanim:"cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilen ters teras çatılarda) ısı yalıtımı",birim:"m²",birimFiyat:640.31,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:840,no:"15.341.2066",tanim:"cm kalınlıkta yüzeyi düzgün XPS levhalar ile (üzeri gezilen ters teras çatılarda) ısı yalıtımı",birim:"m²",birimFiyat:600.94,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:841,no:"15.341.2081",tanim:"basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı",birim:"m²",birimFiyat:466.25,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:842,no:"15.341.2082",tanim:"basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı",birim:"m²",birimFiyat:454.44,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:843,no:"15.341.2083",tanim:"basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı",birim:"m²",birimFiyat:438.69,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:844,no:"15.341.2084",tanim:"basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı",birim:"m²",birimFiyat:600.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:845,no:"15.341.2085",tanim:"basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı",birim:"m²",birimFiyat:572.56,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:846,no:"15.341.2086",tanim:"basınç dayanımlı) ile toprak temaslı temel perde ve duvarlarında su yalıtımı üzerine ısı",birim:"m²",birimFiyat:552.88,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:847,no:"15.341.2101",tanim:"",birim:"m²",birimFiyat:1186.28,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:848,no:"15.341.3001",tanim:"",birim:"m²",birimFiyat:1244.46,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:849,no:"15.341.3002",tanim:"",birim:"m²",birimFiyat:1296.96,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:850,no:"15.341.3003",tanim:"",birim:"m²",birimFiyat:1401.96,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:851,no:"15.341.3004",tanim:"",birim:"m²",birimFiyat:1509.21,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:852,no:"15.341.3021",tanim:"",birim:"m²",birimFiyat:646.88,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:853,no:"15.341.3022",tanim:"",birim:"m²",birimFiyat:666.56,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:854,no:"15.341.3023",tanim:"",birim:"m²",birimFiyat:765.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:855,no:"15.341.3024",tanim:"",birim:"m²",birimFiyat:791.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:856,no:"15.341.3025",tanim:"",birim:"m²",birimFiyat:870.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:857,no:"15.341.3041",tanim:"yada arakat döşeme betonu üzerinde yüke maruz kalan yüzer döşeme uygulaması vb.) ısı",birim:"m²",birimFiyat:240.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:858,no:"15.341.3042",tanim:"yada arakat döşeme betonu üzerinde yüke maruz kalan yüzer döşeme uygulaması vb.) ısı",birim:"m²",birimFiyat:259.69,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:859,no:"15.341.3061",tanim:"(yüke maruz kalmayan uygulamalarda, yükseltilmiş döşeme uygulaması vb.) ısı yalıtımı",birim:"m²",birimFiyat:130.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:860,no:"15.341.3062",tanim:"(yüke maruz kalmayan uygulamalarda, yükseltilmiş döşeme uygulaması vb.) ısı yalıtımı",birim:"m²",birimFiyat:115.45,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:861,no:"15.341.3081",tanim:"döşemelerde ısı ve ses yalıtımı yapılması (Ses yutma katsayısı (ortalama) ≥ 0,80 olan",birim:"m²",birimFiyat:200.63,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:862,no:"15.341.3082",tanim:"döşemelerde ısı ve ses yalıtımı yapılması (Ses yutma katsayısı (ortalama) ≥ 0,90 olan",birim:"m²",birimFiyat:263.63,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:863,no:"15.341.3083",tanim:"döşemelerde ısı ve ses yalıtımı yapılması (Ses yutma katsayısı (ortalama) ≥ 1,00 olan",birim:"m²",birimFiyat:318.75,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:864,no:"15.341.4001",tanim:"8 cm kalınlıkta kaplamasız camyünü şilte ve üzerine su buharı geçişine açık su yalıtım örtüsü",birim:"m²",birimFiyat:289.5,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:865,no:"15.341.4002",tanim:"8 cm kalınlıkta kaplamasız camyünü şilte ve üzerine su buharı geçişine açık su yalıtım örtüsü",birim:"m²",birimFiyat:245.93,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:866,no:"15.341.4003",tanim:"8 cm kalınlıkta bir yüzü cam tülü kaplamalı camyünü şilte ve üzerine su buharı geçişine açık",birim:"m²",birimFiyat:326.25,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:867,no:"15.341.4004",tanim:"8 cm kalınlıkta bir yüzü camtülü kaplamalı camyünü şilte ve üzerine su buharı geçişine açık",birim:"m²",birimFiyat:273.75,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:868,no:"15.341.4005",tanim:"16 cm kalınlıkta kaplamasız camyünü şilte ve üzerine su buharı geçişine açık su yalıtım",birim:"m²",birimFiyat:410.25,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:869,no:"15.341.4006",tanim:"16 cm kalınlıkta kaplamasız camyünü şilte ve üzerine su buharı geçişine açık su yalıtım",birim:"m²",birimFiyat:323.1,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:870,no:"15.341.4007",tanim:"16 cm kalınlıkta bir yüzü camtülü kaplamalı camyünü şilte ve üzerine su buharı geçişine açık",birim:"m²",birimFiyat:464.06,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:871,no:"15.341.4008",tanim:"16 cm kalınlıkta bir yüzü camtülü kaplamalı camyünü şilte ve üzerine su buharı geçişine açık",birim:"m²",birimFiyat:365.63,idare:"Çevre ve Şehircilik",kategori:"Tesisat",kaynak:"2026"},
  {id:872,no:"15.341.4021",tanim:"",birim:"m²",birimFiyat:173.06,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:873,no:"15.341.4022",tanim:"",birim:"m²",birimFiyat:180.94,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:874,no:"15.341.4023",tanim:"",birim:"m²",birimFiyat:184.88,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:875,no:"15.341.4041",tanim:"",birim:"m²",birimFiyat:385.69,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:876,no:"15.341.4042",tanim:"",birim:"m²",birimFiyat:401.44,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:877,no:"15.341.4043",tanim:"",birim:"m²",birimFiyat:417.19,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:878,no:"15.341.5001",tanim:"",birim:"m²",birimFiyat:153.76,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:879,no:"15.345.1001",tanim:"",birim:"m²",birimFiyat:1539.24,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:880,no:"15.345.1002",tanim:"",birim:"m²",birimFiyat:1631.11,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:881,no:"15.345.1003",tanim:"",birim:"m²",birimFiyat:1722.99,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:882,no:"15.345.1004",tanim:"",birim:"m²",birimFiyat:1814.86,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:883,no:"15.345.1005",tanim:"",birim:"m²",birimFiyat:1906.74,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:884,no:"15.345.1006",tanim:"",birim:"m²",birimFiyat:1998.61,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:885,no:"15.345.1101",tanim:"",birim:"m²",birimFiyat:905.08,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:886,no:"15.345.1102",tanim:"",birim:"m²",birimFiyat:996.95,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:887,no:"15.345.1103",tanim:"",birim:"m²",birimFiyat:1088.83,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:888,no:"15.345.1104",tanim:"",birim:"m²",birimFiyat:1180.7,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:889,no:"15.345.1105",tanim:"",birim:"m²",birimFiyat:1272.58,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:890,no:"15.345.1106",tanim:"",birim:"m²",birimFiyat:1364.45,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:891,no:"15.360.1001",tanim:"Alüminyum köse profilinin (fileli) temini ve yerine tesbiti",birim:"m",birimFiyat:50.56,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:892,no:"15.360.1002",tanim:"PVC köse profilinin (fileli) temini ve yerine tespiti",birim:"m",birimFiyat:44.66,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:893,no:"15.360.1003",tanim:"Alüminyum damlalıklı köşe profilinin (fileli) temini ve yerine tespiti",birim:"m",birimFiyat:68.94,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:894,no:"15.360.1004",tanim:"PVC damlalıklı köşe profilinin (fileli) temini ve yerine tespiti",birim:"m",birimFiyat:51.88,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:895,no:"15.360.1005",tanim:"",birim:"m",birimFiyat:98.75,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:896,no:"15.360.1006",tanim:"",birim:"m",birimFiyat:196.25,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:897,no:"15.360.1007",tanim:"",birim:"m",birimFiyat:74.19,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:898,no:"15.365.1001",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta PVC esaslı yer döseme malzemeleri ile döseme",birim:"m²",birimFiyat:1047.03,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:899,no:"15.365.1002",tanim:"tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı yer döseme kaplaması yapılması",birim:"m²",birimFiyat:948.59,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:900,no:"15.365.1003",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme",birim:"m²",birimFiyat:1027.34,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:901,no:"15.365.1004",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta PVC esaslı karo yer döseme malzemeleri ile",birim:"m²",birimFiyat:1270.15,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:902,no:"15.365.1005",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme",birim:"m²",birimFiyat:1276.71,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:903,no:"15.365.1006",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme",birim:"m²",birimFiyat:1447.34,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:904,no:"15.365.1007",tanim:"tesviyesi yapılması ve üzerine 3 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme",birim:"m²",birimFiyat:1145.46,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:905,no:"15.365.1008",tanim:"tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı yer döseme kaplaması yapılması",birim:"m²",birimFiyat:1191.4,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:906,no:"15.365.1009",tanim:"tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı karo yer döseme kaplaması",birim:"m²",birimFiyat:1394.84,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:907,no:"15.365.1021",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta PVC esaslı yer döseme malzemeleri ile döseme",birim:"m²",birimFiyat:1025.4,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:908,no:"15.365.1022",tanim:"tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı yer döseme kaplaması yapılması",birim:"m²",birimFiyat:926.96,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:909,no:"15.365.1023",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme",birim:"m²",birimFiyat:1005.71,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:910,no:"15.365.1024",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta PVC esaslı karo yer döseme malzemeleri ile",birim:"m²",birimFiyat:1248.53,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:911,no:"15.365.1025",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme",birim:"m²",birimFiyat:1255.09,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:912,no:"15.365.1026",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme",birim:"m²",birimFiyat:1425.71,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:913,no:"15.365.1027",tanim:"tesviyesi yapılması ve üzerine 3 mm kalınlıkta Pvc esaslı yer döseme malzemeleri ile döseme",birim:"m²",birimFiyat:1123.84,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:914,no:"15.365.1028",tanim:"tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı yer döseme kaplaması yapılması",birim:"m²",birimFiyat:1169.78,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:915,no:"15.365.1029",tanim:"tesviyesi yapılması ve üzerine 2mm kalınlıkda PVC esaslı karo yer döseme kaplaması",birim:"m²",birimFiyat:1373.21,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:916,no:"15.365.1101",tanim:"tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde",birim:"m²",birimFiyat:2087.65,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:917,no:"15.365.1102",tanim:"tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde",birim:"m²",birimFiyat:2612.65,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:918,no:"15.365.1103",tanim:"tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde",birim:"m²",birimFiyat:3268.9,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:919,no:"15.365.1111",tanim:"tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde",birim:"m²",birimFiyat:2066.03,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:920,no:"15.365.1112",tanim:"tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde",birim:"m²",birimFiyat:2591.03,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:921,no:"15.365.1113",tanim:"tesviyesi yapılması ve üzerine PVC esaslı spor zemin malzemeleri ile kapalı spor zeminlerde",birim:"m²",birimFiyat:3247.28,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:922,no:"15.365.1501",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf",birim:"m²",birimFiyat:1192.09,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:923,no:"15.365.1502",tanim:"tesviyesi yapılması ve üzerine 2,5 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf",birim:"m²",birimFiyat:1119.9,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:924,no:"15.365.1503",tanim:"tesviyesi yapılması ve üzerine 3,2 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf",birim:"m²",birimFiyat:1448.03,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:925,no:"15.365.1511",tanim:"tesviyesi yapılması ve üzerine 2 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf",birim:"m²",birimFiyat:1170.46,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:926,no:"15.365.1512",tanim:"tesviyesi yapılması ve üzerine 2,5 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf",birim:"m²",birimFiyat:1098.28,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:927,no:"15.365.1513",tanim:"tesviyesi yapılması ve üzerine 3,2 mm kalınlıkta Linolyum zemin kaplaması yapılması (Sınıf",birim:"m²",birimFiyat:1426.4,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:928,no:"15.365.1701",tanim:"PVC esaslı esnek süpürgelik temini ve yerine tespit edilmesi.",birim:"m",birimFiyat:72.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:929,no:"15.365.1702",tanim:"PVC esaslı kendinden dönüşlü kepli süpürgelik temini ve yerine tespit edilmesi.",birim:"m",birimFiyat:98.44,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:930,no:"15.365.1751",tanim:"PVC esaslı geçiş profili (4 cm genişliğinde) temini ve yerine monte edilmesi.",birim:"m",birimFiyat:178.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:931,no:"15.365.1752",tanim:"Alüminyum esaslı geçiş profili (4 cm genişliğinde) temini ve yerine monte edilmesi.",birim:"m",birimFiyat:249.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:932,no:"15.375.1002",tanim:"I.kalite, beyaz seramik yer karoları ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:628.7,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:933,no:"15.375.1003",tanim:"",birim:"m²",birimFiyat:684.81,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:934,no:"15.375.1004",tanim:"I.kalite, beyaz seramik yer karoları ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:691.44,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:935,no:"15.375.1052",tanim:"I.kalite, renkli seramik yer karoları ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:639.2,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:936,no:"15.375.1053",tanim:"",birim:"m²",birimFiyat:704.69,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:937,no:"15.375.1054",tanim:"I.kalite, renkli seramik yer karoları ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:704.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:938,no:"15.380.1003",tanim:"I.kalite, beyaz seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması yapılması",birim:"m²",birimFiyat:714.7,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:939,no:"15.380.1005",tanim:"I.kalite, beyaz seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması yapılması",birim:"m²",birimFiyat:692.39,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:940,no:"15.380.1006",tanim:"özelliğinde, I.kalite, beyaz seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması",birim:"m²",birimFiyat:816.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:941,no:"15.380.1007",tanim:"türlü desen ve yüzey özelliğinde, I.kalite, beyaz seramik duvar karoları ile 3 mm derz aralıklı",birim:"m²",birimFiyat:725.33,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:942,no:"15.380.1053",tanim:"I.kalite, renkli seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması yapılması",birim:"m²",birimFiyat:734.39,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:943,no:"15.380.1055",tanim:"I.kalite, renkli seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması yapılması",birim:"m²",birimFiyat:700.26,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:944,no:"15.380.1056",tanim:"özelliğinde, I.kalite, renkli seramik duvar karoları ile 3 mm derz aralıklı duvar kaplaması",birim:"m²",birimFiyat:833.54,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:945,no:"15.380.1057",tanim:"türlü desen ve yüzey özelliğinde, I.kalite, renkli seramik duvar karoları ile 3 mm derz",birim:"m²",birimFiyat:743.7,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:946,no:"15.385.1004",tanim:"I.kalite, beyaz, sırlı porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:787.83,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:947,no:"15.385.1005",tanim:"",birim:"m²",birimFiyat:809.54,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:948,no:"15.385.1006",tanim:"I.kalite, beyaz, sırlı porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:809.54,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:949,no:"15.385.1008",tanim:"",birim:"m²",birimFiyat:883.74,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:950,no:"15.385.1009",tanim:"",birim:"m²",birimFiyat:959.26,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:951,no:"15.385.1010",tanim:"",birim:"m²",birimFiyat:895.66,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:952,no:"15.385.1024",tanim:"I.kalite, renkli, sırlı porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:804.89,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:953,no:"15.385.1025",tanim:"",birim:"m²",birimFiyat:829.41,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:954,no:"15.385.1026",tanim:"I.kalite, renkli, sırlı porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:829.41,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:955,no:"15.385.1028",tanim:"",birim:"m²",birimFiyat:895.66,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:956,no:"15.385.1029",tanim:"",birim:"m²",birimFiyat:971.19,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:957,no:"15.385.1030",tanim:"",birim:"m²",birimFiyat:923.49,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:958,no:"15.385.1043",tanim:"porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı",birim:"m²",birimFiyat:995.89,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:959,no:"15.385.1044",tanim:"I.kalite, beyaz, sırlı porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması",birim:"m²",birimFiyat:826.58,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:960,no:"15.385.1045",tanim:"porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı",birim:"m²",birimFiyat:848.29,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:961,no:"15.385.1046",tanim:"I.kalite, beyaz, sırlı porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması",birim:"m²",birimFiyat:848.29,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:962,no:"15.385.1049",tanim:"porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı",birim:"m²",birimFiyat:998.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:963,no:"15.385.1050",tanim:"porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı",birim:"m²",birimFiyat:934.41,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:964,no:"15.385.1063",tanim:"porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı",birim:"m²",birimFiyat:894.83,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:965,no:"15.385.1064",tanim:"I.kalite, renkli, sırlı porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması",birim:"m²",birimFiyat:843.64,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:966,no:"15.385.1065",tanim:"porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı",birim:"m²",birimFiyat:868.16,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:967,no:"15.385.1066",tanim:"I.kalite, renkli, sırlı porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması",birim:"m²",birimFiyat:868.16,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:968,no:"15.385.1069",tanim:"porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı",birim:"m²",birimFiyat:1009.94,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:969,no:"15.385.1070",tanim:"porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı",birim:"m²",birimFiyat:962.24,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:970,no:"15.390.1004",tanim:"I.kalite, mat, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:833.76,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:971,no:"15.390.1005",tanim:"",birim:"m²",birimFiyat:870.49,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:972,no:"15.390.1006",tanim:"yüzey özelliğinde, I.kalite, mat, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması",birim:"m²",birimFiyat:955.29,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:973,no:"15.390.1008",tanim:"mat, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:1034.79,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:974,no:"15.390.1009",tanim:"mat, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:1034.79,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:975,no:"15.390.1010",tanim:"mat, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:1042.74,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:976,no:"15.390.1024",tanim:"özelliğinde, I.kalite, parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması",birim:"m²",birimFiyat:954.51,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:977,no:"15.390.1025",tanim:"parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:997.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:978,no:"15.390.1026",tanim:"yüzey özelliğinde, I.kalite, parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme",birim:"m²",birimFiyat:1111.64,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:979,no:"15.390.1028",tanim:"parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:1189.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:980,no:"15.390.1029",tanim:"parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:1199.09,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:981,no:"15.390.1030",tanim:"parlak, sırsız porselen karo ile 3 mm derz aralıklı döşeme kaplaması yapılması (karo",birim:"m²",birimFiyat:1240.16,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:982,no:"15.390.1043",tanim:"porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı",birim:"m²",birimFiyat:923.7,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:983,no:"15.390.1044",tanim:"I.kalite, mat, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması",birim:"m²",birimFiyat:872.51,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:984,no:"15.390.1045",tanim:"porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo yapıştırıcısı",birim:"m²",birimFiyat:909.24,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:985,no:"15.390.1046",tanim:"yüzey özelliğinde, I.kalite, mat, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe",birim:"m²",birimFiyat:994.04,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:986,no:"15.390.1049",tanim:"mat, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo",birim:"m²",birimFiyat:1073.54,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:987,no:"15.390.1050",tanim:"mat, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo",birim:"m²",birimFiyat:1081.49,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:988,no:"15.390.1063",tanim:"parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo",birim:"m²",birimFiyat:1044.45,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:989,no:"15.390.1064",tanim:"özelliğinde, I.kalite, parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe",birim:"m²",birimFiyat:993.26,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:990,no:"15.390.1065",tanim:"parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo",birim:"m²",birimFiyat:1036.44,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:991,no:"15.390.1066",tanim:"yüzey özelliğinde, I.kalite, parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe",birim:"m²",birimFiyat:1150.39,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:992,no:"15.390.1069",tanim:"parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo",birim:"m²",birimFiyat:1237.84,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:993,no:"15.390.1070",tanim:"parlak, sırsız porselen karo ile 3 mm derz aralıklı duvar ve cephe kaplaması yapılması (karo",birim:"m²",birimFiyat:1278.91,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:994,no:"15.400.1011",tanim:"",birim:"m²",birimFiyat:1692.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:995,no:"15.400.1012",tanim:"",birim:"m²",birimFiyat:1767.06,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:996,no:"15.400.1013",tanim:"Şartları (Sınıf 2) Yüzey alanı <= 1100cm² ebatlarda ve kırılma dayanımı > 2,5 kN, honlu",birim:"m²",birimFiyat:1744.75,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:997,no:"15.400.1014",tanim:"Şartları (Sınıf 3) 1100 cm²< Yüzey alanı < 1800 cm² ebatlarda ve kırılma dayanımı > 3 kN,",birim:"m²",birimFiyat:1784.13,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:998,no:"15.400.1015",tanim:"Şartları (Sınıf 3) Yüzey alanı >= 1800cm² ebatlarda ve kırılma dayanımı > 3 kN, honlu veya",birim:"m²",birimFiyat:2007.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:999,no:"15.400.1111",tanim:"",birim:"m²",birimFiyat:1915.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1000,no:"15.400.1112",tanim:"",birim:"m²",birimFiyat:1969.19,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1001,no:"15.400.1113",tanim:"Şartları (Sınıf 2) Yüzey alanı <= 1100cm² ebatlarda ve kırılma dayanımı > 2,5 kN, honlu",birim:"m²",birimFiyat:1967.88,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1002,no:"15.400.1114",tanim:"Şartları (Sınıf 3) 1100 cm²< Yüzey alanı < 1800 cm² ebatlarda ve kırılma dayanımı > 3 kN,",birim:"m²",birimFiyat:2007.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1003,no:"15.400.1115",tanim:"Şartları (Sınıf 3) Yüzey alanı >= 1800cm² ebatlarda ve kırılma dayanımı > 3 kN, honlu veya",birim:"m²",birimFiyat:2156.88,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1004,no:"15.400.1211",tanim:"",birim:"m²",birimFiyat:1915.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1005,no:"15.400.1212",tanim:"",birim:"m²",birimFiyat:1969.19,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1006,no:"15.400.1213",tanim:"(Kırılma Yükü Şartları (Sınıf 2) Yüzey alanı <= 1100cm² ebatlarda ve kırılma dayanımı >",birim:"m²",birimFiyat:1967.88,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1007,no:"15.400.1214",tanim:"(Kırılma Yükü Şartları (Sınıf 3) 1100 cm²< Yüzey alanı < 1800 cm² ebatlarda ve kırılma",birim:"m²",birimFiyat:2007.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1008,no:"15.400.1215",tanim:"(Kırılma Yükü Şartları (Sınıf 3) Yüzey alanı >= 1800cm² ebatlarda ve kırılma dayanımı > 3",birim:"m²",birimFiyat:2156.88,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1009,no:"15.400.1311",tanim:"",birim:"m²",birimFiyat:2680.56,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1010,no:"15.400.1312",tanim:"",birim:"m²",birimFiyat:2772.44,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1011,no:"15.400.1313",tanim:"Şartları (Sınıf 2) Yüzey alanı <= 1100cm² ebatlarda ve kırılma dayanımı > 2,5 kN, honlu",birim:"m²",birimFiyat:2772.44,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1012,no:"15.400.1314",tanim:"Şartları (Sınıf 3) 1100 cm²< Yüzey alanı < 1800 cm² ebatlarda ve kırılma dayanımı > 3 kN,",birim:"m²",birimFiyat:2845.94,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1013,no:"15.400.1315",tanim:"Şartları (Sınıf 3) Yüzey alanı >= 1800cm² ebatlarda ve kırılma dayanımı > 3 kN, honlu veya",birim:"m²",birimFiyat:3015.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1014,no:"15.405.1011",tanim:"Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa, Aşınma direnç sınıfı (2-G), Yüzey",birim:"m²",birimFiyat:1671.25,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1015,no:"15.405.1012",tanim:"Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), 1600 <",birim:"m²",birimFiyat:1767.06,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1016,no:"15.405.1013",tanim:"Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H), Yüzey alanı",birim:"m²",birimFiyat:1744.75,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1017,no:"15.405.1014",tanim:"Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),1600 <",birim:"m²",birimFiyat:1877.31,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1018,no:"15.405.1015",tanim:"Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), Yüzey alanı",birim:"m²",birimFiyat:1858.94,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1019,no:"15.405.1016",tanim:"Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), 1600 <",birim:"m²",birimFiyat:1995.44,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1020,no:"15.405.1111",tanim:"Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), Yüzey alanı",birim:"m²",birimFiyat:1744.75,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1021,no:"15.405.1112",tanim:"Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), 1600 <",birim:"m²",birimFiyat:1840.56,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1022,no:"15.405.1113",tanim:"Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H), Yüzey alanı",birim:"m²",birimFiyat:1840.56,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1023,no:"15.405.1114",tanim:"Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),1600 <",birim:"m²",birimFiyat:1988.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1024,no:"15.405.1115",tanim:"Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), Yüzey alanı",birim:"m²",birimFiyat:1969.19,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1025,no:"15.405.1116",tanim:"Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), 1600 <",birim:"m²",birimFiyat:2083.38,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1026,no:"15.405.1211",tanim:"Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), Yüzey alanı",birim:"m²",birimFiyat:1877.31,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1027,no:"15.405.1212",tanim:"Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), 1600 <",birim:"m²",birimFiyat:1988.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1028,no:"15.405.1213",tanim:"Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H), Yüzey alanı",birim:"m²",birimFiyat:1988.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1029,no:"15.405.1214",tanim:"Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),1600 <",birim:"m²",birimFiyat:2117.5,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1030,no:"15.405.1215",tanim:"Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), Yüzey alanı",birim:"m²",birimFiyat:2083.38,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1031,no:"15.405.1216",tanim:"Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), 1600 <",birim:"m²",birimFiyat:2231.69,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1032,no:"15.405.1311",tanim:"Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), Yüzey alanı",birim:"m²",birimFiyat:1840.56,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1033,no:"15.405.1312",tanim:"Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), 1600 <",birim:"m²",birimFiyat:1969.19,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1034,no:"15.405.1313",tanim:"Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H), Yüzey alanı",birim:"m²",birimFiyat:1895.69,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1035,no:"15.405.1314",tanim:"Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),1600 <",birim:"m²",birimFiyat:2044.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1036,no:"15.405.1315",tanim:"Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), Yüzey alanı",birim:"m²",birimFiyat:2019.06,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1037,no:"15.405.1316",tanim:"Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), 1600 <",birim:"m²",birimFiyat:2158.19,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1038,no:"15.405.1411",tanim:"Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), Yüzey alanı",birim:"m²",birimFiyat:1784.13,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1039,no:"15.405.1412",tanim:"Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G), 1600 <",birim:"m²",birimFiyat:1877.31,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1040,no:"15.405.1413",tanim:"Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H), Yüzey alanı",birim:"m²",birimFiyat:1895.69,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1041,no:"15.405.1414",tanim:"Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),1600 cm² <",birim:"m²",birimFiyat:2044.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1042,no:"15.405.1415",tanim:"Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), Yüzey alanı",birim:"m²",birimFiyat:1995.44,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1043,no:"15.405.1416",tanim:"Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I), 1600 <",birim:"m²",birimFiyat:2135.88,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1044,no:"15.405.1511",tanim:"Dayanımı Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G),",birim:"m²",birimFiyat:2065.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1045,no:"15.405.1512",tanim:"Dayanımı Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G),",birim:"m²",birimFiyat:2175.25,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1046,no:"15.405.1513",tanim:"Dayanımı Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),",birim:"m²",birimFiyat:2175.25,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1047,no:"15.405.1514",tanim:"Dayanımı Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı",birim:"m²",birimFiyat:2250.06,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1048,no:"15.405.1515",tanim:"Dayanımı Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I),",birim:"m²",birimFiyat:2250.06,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1049,no:"15.405.1516",tanim:"Dayanımı Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I),",birim:"m²",birimFiyat:2341.94,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1050,no:"15.405.1611",tanim:"Yükü Şartları (Sınıf 1) Eğilme dayanımı minimum 2,8 Mpa Aşınma direnç sınıfı (2-G),",birim:"m²",birimFiyat:1729.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1051,no:"15.405.1612",tanim:"Yükü Sartları (Sınıf 1) Egilme dayanımı minimum 2,8 Mpa Asınma direnç sınıfı (2-G), 1600",birim:"m²",birimFiyat:1840.56,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1052,no:"15.405.1613",tanim:"Dayanımı Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı (3-H),",birim:"m²",birimFiyat:1969.19,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1053,no:"15.405.1614",tanim:"Dayanımı Şartları (Sınıf 2) Eğilme dayanımı minimum 3,2 Mpa Aşınma direnç sınıfı",birim:"m²",birimFiyat:2025.63,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1054,no:"15.405.1615",tanim:"Dayanımı Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I),",birim:"m²",birimFiyat:2007.25,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1055,no:"15.405.1616",tanim:"Dayanımı Şartları (Sınıf 3) Eğilme dayanımı minimum 4,0 Mpa Aşınma direnç sınıfı (4-I),",birim:"m²",birimFiyat:2117.5,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1056,no:"15.405.1701",tanim:"",birim:"m",birimFiyat:289.01,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1057,no:"15.410.1011",tanim:"",birim:"m²",birimFiyat:2217.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1058,no:"15.410.1012",tanim:"",birim:"m²",birimFiyat:2437.5,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1059,no:"15.410.1013",tanim:"",birim:"m²",birimFiyat:2360.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1060,no:"15.410.1015",tanim:"",birim:"m²",birimFiyat:2464.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1061,no:"15.410.1017",tanim:"",birim:"m²",birimFiyat:2548.5,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1062,no:"15.410.1111",tanim:"",birim:"m²",birimFiyat:2370.56,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1063,no:"15.410.1112",tanim:"",birim:"m²",birimFiyat:2591.06,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1064,no:"15.410.1113",tanim:"",birim:"m²",birimFiyat:2545.74,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1065,no:"15.410.1115",tanim:"",birim:"m²",birimFiyat:2673.14,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1066,no:"15.410.1117",tanim:"",birim:"m²",birimFiyat:2776.65,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1067,no:"15.410.1211",tanim:"",birim:"m²",birimFiyat:2523.9,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1068,no:"15.410.1213",tanim:"",birim:"m²",birimFiyat:2677.46,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1069,no:"15.410.1311",tanim:"",birim:"m",birimFiyat:1400.11,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1070,no:"15.410.1313",tanim:"",birim:"m",birimFiyat:1489.85,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1071,no:"15.410.1411",tanim:"",birim:"m²",birimFiyat:3729.4,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1072,no:"15.410.1413",tanim:"",birim:"m²",birimFiyat:3915.14,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1073,no:"15.410.1511",tanim:"",birim:"m²",birimFiyat:3923.15,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1074,no:"15.410.1513",tanim:"",birim:"m²",birimFiyat:4108.89,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1075,no:"15.410.1611",tanim:"",birim:"m²",birimFiyat:4127.53,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1076,no:"15.410.1613",tanim:"",birim:"m²",birimFiyat:4313.26,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1077,no:"15.410.1711",tanim:"",birim:"m²",birimFiyat:3427.03,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1078,no:"15.410.1713",tanim:"",birim:"m²",birimFiyat:3580.59,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1079,no:"15.415.1011",tanim:"",birim:"m²",birimFiyat:2433.56,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1080,no:"15.415.1012",tanim:"",birim:"m²",birimFiyat:2654.06,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1081,no:"15.415.1013",tanim:"",birim:"m²",birimFiyat:2621.94,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1082,no:"15.415.1015",tanim:"",birim:"m²",birimFiyat:2758.94,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1083,no:"15.415.1017",tanim:"",birim:"m²",birimFiyat:2870.25,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1084,no:"15.415.1111",tanim:"",birim:"m²",birimFiyat:2259.0,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1085,no:"15.415.1112",tanim:"",birim:"m²",birimFiyat:2479.5,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1086,no:"15.415.1113",tanim:"",birim:"m²",birimFiyat:2410.8,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1087,no:"15.415.1115",tanim:"",birim:"m²",birimFiyat:2521.2,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1088,no:"15.415.1117",tanim:"",birim:"m²",birimFiyat:2610.9,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1089,no:"15.415.1211",tanim:"",birim:"m²",birimFiyat:2740.46,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1090,no:"15.415.1213",tanim:"",birim:"m²",birimFiyat:2565.9,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1091,no:"15.415.1311",tanim:"",birim:"m",birimFiyat:1526.68,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1092,no:"15.415.1313",tanim:"",birim:"m",birimFiyat:1424.66,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1093,no:"15.415.1411",tanim:"",birim:"m²",birimFiyat:3991.34,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1094,no:"15.415.1413",tanim:"",birim:"m²",birimFiyat:3780.2,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1095,no:"15.415.1511",tanim:"",birim:"m²",birimFiyat:4185.09,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1096,no:"15.415.1513",tanim:"",birim:"m²",birimFiyat:3973.95,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1097,no:"15.415.1611",tanim:"",birim:"m²",birimFiyat:4389.46,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1098,no:"15.415.1613",tanim:"",birim:"m²",birimFiyat:4178.33,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1099,no:"15.415.1711",tanim:"",birim:"m²",birimFiyat:3643.59,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1100,no:"15.415.1713",tanim:"",birim:"m²",birimFiyat:3469.03,idare:"Çevre ve Şehircilik",kategori:"Kaba İnşaat",kaynak:"2026"},
  {id:1101,no:"15.420.1011",tanim:"4 cm kalınlığında andezit levha ile döşeme kaplaması yapılması (30cmxserbest boy)",birim:"m²",birimFiyat:2414.94,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1102,no:"15.420.1012",tanim:"",birim:"m²",birimFiyat:2996.19,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1103,no:"15.420.1111",tanim:"3 cm kalınlığında andezit levha ile duvar kaplaması yapılması (30cmxserbest boy)",birim:"m²",birimFiyat:2584.28,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1104,no:"15.420.1211",tanim:"3 cm kalınlığında andezit levha ile söve yapılması",birim:"m²",birimFiyat:2621.68,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1105,no:"15.430.1011",tanim:"",birim:"m",birimFiyat:2254.3,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1106,no:"15.430.1012",tanim:"",birim:"m",birimFiyat:2358.98,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1107,no:"15.430.1013",tanim:"",birim:"m",birimFiyat:2358.98,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1108,no:"15.430.1014",tanim:"",birim:"m",birimFiyat:2482.2,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1109,no:"15.430.1015",tanim:"",birim:"m",birimFiyat:2859.16,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1110,no:"15.430.1111",tanim:"",birim:"m",birimFiyat:2407.34,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1111,no:"15.430.1112",tanim:"",birim:"m",birimFiyat:2543.81,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1112,no:"15.430.1113",tanim:"",birim:"m",birimFiyat:2543.81,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1113,no:"15.430.1114",tanim:"",birim:"m",birimFiyat:2622.65,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1114,no:"15.430.1115",tanim:"",birim:"m",birimFiyat:3040.69,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1115,no:"15.430.1201",tanim:"",birim:"m",birimFiyat:394.01,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1116,no:"15.430.1202",tanim:"",birim:"m",birimFiyat:412.39,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1117,no:"15.430.1311",tanim:"",birim:"m²",birimFiyat:3481.44,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1118,no:"15.430.1312",tanim:"",birim:"m²",birimFiyat:3691.44,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1119,no:"15.430.1313",tanim:"",birim:"m²",birimFiyat:3793.81,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1120,no:"15.430.1314",tanim:"",birim:"m²",birimFiyat:4349.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1121,no:"15.430.1411",tanim:"",birim:"m²",birimFiyat:3619.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1122,no:"15.430.1412",tanim:"",birim:"m²",birimFiyat:3724.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1123,no:"15.430.1413",tanim:"",birim:"m²",birimFiyat:4053.69,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1124,no:"15.430.1414",tanim:"",birim:"m²",birimFiyat:4349.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1125,no:"15.430.1511",tanim:"",birim:"m²",birimFiyat:4055.0,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1126,no:"15.430.1512",tanim:"",birim:"m²",birimFiyat:4279.44,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1127,no:"15.430.1513",tanim:"",birim:"m²",birimFiyat:4574.75,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1128,no:"15.430.1514",tanim:"",birim:"m²",birimFiyat:4782.13,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1129,no:"15.435.1001",tanim:"",birim:"m²",birimFiyat:839.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1130,no:"15.435.1002",tanim:"",birim:"m²",birimFiyat:865.56,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1131,no:"15.435.1003",tanim:"",birim:"m²",birimFiyat:891.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1132,no:"15.435.1004",tanim:"",birim:"m²",birimFiyat:813.06,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1133,no:"15.435.1005",tanim:"",birim:"m²",birimFiyat:839.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1134,no:"15.435.1006",tanim:"",birim:"m²",birimFiyat:865.56,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1135,no:"15.435.1101",tanim:"",birim:"m²",birimFiyat:944.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1136,no:"15.435.1102",tanim:"",birim:"m²",birimFiyat:983.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1137,no:"15.435.1103",tanim:"",birim:"m²",birimFiyat:911.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1138,no:"15.435.1104",tanim:"",birim:"m²",birimFiyat:950.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1139,no:"15.435.1211",tanim:"",birim:"m",birimFiyat:308.04,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1140,no:"15.435.1212",tanim:"",birim:"m",birimFiyat:353.98,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1141,no:"15.435.1213",tanim:"",birim:"m",birimFiyat:294.91,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1142,no:"15.435.1214",tanim:"",birim:"m",birimFiyat:334.29,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1143,no:"15.435.1215",tanim:"10x15x50 cm boyutlarında andezit bordür temini ve yerine döşenmesi",birim:"m",birimFiyat:767.41,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1144,no:"15.435.1216",tanim:"10x20x50 cm boyutlarında andezit bordür temini ve yerine döşenmesi",birim:"m",birimFiyat:830.41,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1145,no:"15.435.1217",tanim:"10x20x70 cm boyutlarında andezit bordür temini ve yerine döşenmesi",birim:"m",birimFiyat:838.66,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1146,no:"15.435.1311",tanim:"",birim:"m",birimFiyat:406.48,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1147,no:"15.435.1312",tanim:"",birim:"m",birimFiyat:386.79,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1148,no:"15.435.1313",tanim:"50 x 20 cm boyutlarında andezit oluk tası döşenmesi",birim:"m",birimFiyat:957.73,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1149,no:"15.435.7001",tanim:"",birim:"m²",birimFiyat:935.06,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1150,no:"15.435.7002",tanim:"",birim:"m²",birimFiyat:986.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1151,no:"15.435.7003",tanim:"",birim:"m²",birimFiyat:1123.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1152,no:"15.440.1001",tanim:"duvar, tavan ve cephelerde kaplama üstü dilatasyon fugası yapılması (50 mm genislikte",birim:"m",birimFiyat:403.41,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1153,no:"15.440.1002",tanim:"mm hareket kapasiteli, profil yüksekligi min. 13mm, kanat genisligi min.45 mm) duvar ve",birim:"m",birimFiyat:588.2,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1154,no:"15.440.1003",tanim:"zeminlerde kaplama üstü dilatasyon fugası yapılması (50 mm genislikte dilatasyonlar için)",birim:"m",birimFiyat:493.98,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1155,no:"15.440.1004",tanim:"hareket kapasiteli, profil yüksekligi min. 35 mm, kanat genisligi min. 45mm) zeminde",birim:"m",birimFiyat:873.7,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1156,no:"15.440.1005",tanim:"genisligi min. 45mm, fitilin yerlestirildigi mesnetler arası ilave elemanlarla güçlendirilmis)",birim:"m",birimFiyat:1304.2,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1157,no:"15.440.1006",tanim:"mm hareket kapasiteli, profil yüksekligi min. 15 mm, kanat genisligi min. 45mm) zeminde",birim:"m",birimFiyat:816.39,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1158,no:"15.440.1007",tanim:"kanat genisligi min. 45mm, fitilin yerlestirildigi mesnetler arası ilave elemanlarla",birim:"m",birimFiyat:913.51,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1159,no:"15.440.1008",tanim:"",birim:"m",birimFiyat:693.94,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1160,no:"15.445.1001",tanim:"",birim:"m²",birimFiyat:1381.39,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1161,no:"15.445.1002",tanim:"",birim:"m²",birimFiyat:1638.89,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1162,no:"15.445.1003",tanim:"",birim:"m²",birimFiyat:2218.26,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1163,no:"15.445.1004",tanim:"",birim:"m²",birimFiyat:2411.39,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1164,no:"15.450.1011",tanim:"Mozayik denizlik yapılması (normal çimentolu)",birim:"m²",birimFiyat:5593.55,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1165,no:"15.450.1012",tanim:"Mozayik denizlik yapılması (beyaz çimentolu)",birim:"m²",birimFiyat:5655.7,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1166,no:"15.450.1013",tanim:"Mozayik parapet yapılması (normal çimentolu)",birim:"m²",birimFiyat:5562.39,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1167,no:"15.450.1014",tanim:"Mozaik parapet yapılması (beyaz çimentolu)",birim:"m²",birimFiyat:5614.19,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1168,no:"15.450.1015",tanim:"",birim:"m²",birimFiyat:4096.5,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1169,no:"15.450.1016",tanim:"",birim:"m²",birimFiyat:4148.3,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1170,no:"15.451.1001",tanim:"",birim:"kg",birimFiyat:284.06,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1171,no:"15.451.1002",tanim:"Kaplama Profili İle Binaların Dış Yüzeylerinde (Cephe) Kaplama Yapılması (Yardımcı",birim:"kg",birimFiyat:297.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1172,no:"15.455.1002",tanim:"",birim:"kg",birimFiyat:303.98,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1173,no:"15.455.1003",tanim:"",birim:"kg",birimFiyat:339.98,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1174,no:"15.455.1004",tanim:"konulması (Sert PVC doğrama profillerinden her çeşit kapı, pencere, kaplama ve benzeri",birim:"kg",birimFiyat:267.98,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1175,no:"15.455.1005",tanim:"ve yerine konulması (Sert PVC doğrama profillerinden her çeşit kapı, pencere, kaplama ve",birim:"kg",birimFiyat:291.98,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1176,no:"15.460.1001",tanim:"",birim:"kg",birimFiyat:479.83,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1177,no:"15.460.1002",tanim:"",birim:"kg",birimFiyat:526.64,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1178,no:"15.460.1003",tanim:"",birim:"kg",birimFiyat:479.83,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1179,no:"15.460.1004",tanim:"",birim:"kg",birimFiyat:485.18,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1180,no:"15.460.1005",tanim:"",birim:"kg",birimFiyat:525.3,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1181,no:"15.460.1006",tanim:"",birim:"kg",birimFiyat:533.33,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1182,no:"15.460.1007",tanim:"",birim:"kg",birimFiyat:538.68,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1183,no:"15.460.1008",tanim:"",birim:"kg",birimFiyat:526.64,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1184,no:"15.460.1009",tanim:"",birim:"kg",birimFiyat:538.68,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1185,no:"15.460.1010",tanim:"",birim:"kg",birimFiyat:526.64,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1186,no:"15.465.1001",tanim:"Gömme iç kapı kilidinin yerine takılması",birim:"adet",birimFiyat:137.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1187,no:"15.465.1002",tanim:"Gömme iç kapı kilidinin yerine takılması",birim:"adet",birimFiyat:137.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1188,no:"15.465.1003",tanim:"Gömme makaralı iç kapı kilidinin yerine takılması",birim:"adet",birimFiyat:250.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1189,no:"15.465.1004",tanim:"Gömme silindirli iç ve dış kapı kilidinin yerine takılması",birim:"adet",birimFiyat:375.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1190,no:"15.465.1005",tanim:"Gömme makaralı silindirli iç ve dış kapı kilidinin yerine takılması",birim:"adet",birimFiyat:375.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1191,no:"15.465.1006",tanim:"Gömme makaralı silindirli iç ve dış kapı kilidinin yerine takılması",birim:"adet",birimFiyat:375.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1192,no:"15.465.1007",tanim:"Silindirli traşlı dış kapı kilidinin yerine takılması",birim:"adet",birimFiyat:437.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1193,no:"15.465.1008",tanim:"Kapı kolu ve aynalarının yerine takılması",birim:"adet",birimFiyat:137.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1194,no:"15.465.1009",tanim:"Lastik başlı tamponun yerine takılması",birim:"adet",birimFiyat:23.13,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1195,no:"15.465.1010",tanim:"Menteşenin yerine takılması",birim:"adet",birimFiyat:23.13,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1196,no:"15.465.1011",tanim:"Yaylı menteşenin yerine takılması",birim:"adet",birimFiyat:212.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1197,no:"15.465.1012",tanim:"Sürgünün yerine takılması",birim:"adet",birimFiyat:31.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1198,no:"15.465.1013",tanim:"Stopun yerine takılması",birim:"adet",birimFiyat:112.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1199,no:"15.465.1101",tanim:"İspanyolet takımının yerine takılması",birim:"adet",birimFiyat:125.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1200,no:"15.465.1102",tanim:"Vasistas takımının yerine takılması",birim:"adet",birimFiyat:35.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1201,no:"15.465.1103",tanim:"Vasistas takımının yerine takılması",birim:"adet",birimFiyat:93.75,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1202,no:"15.465.1104",tanim:"Mandalın yerine takılması",birim:"adet",birimFiyat:75.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1203,no:"15.465.1105",tanim:"Sürgünün yerine takılması",birim:"adet",birimFiyat:25.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1204,no:"15.465.1106",tanim:"Lastik başlı tamponun yerine takılması",birim:"adet",birimFiyat:31.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1205,no:"15.465.1107",tanim:"Yaylı tespit mandalının yerine takılması",birim:"adet",birimFiyat:37.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1206,no:"15.465.1108",tanim:"Kontrpua takımının yerine takılması",birim:"kg",birimFiyat:37.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1207,no:"15.465.1109",tanim:"Sürme kanat tutamağının yerine takılması",birim:"adet",birimFiyat:106.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1208,no:"15.465.1110",tanim:"Kavramalı ispanyolet takımının yerine takılması",birim:"adet",birimFiyat:112.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1209,no:"15.465.1111",tanim:"Kavramalı ispanyolet takımının yerine takılması",birim:"adet",birimFiyat:125.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1210,no:"15.465.1112",tanim:"Kavramalı ispanyolet takımının yerine takılması",birim:"adet",birimFiyat:150.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1211,no:"15.465.1113",tanim:"Kavramalı ispanyolet takımının yerine takılması",birim:"adet",birimFiyat:150.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1212,no:"15.465.1114",tanim:"Kavramalı ispanyolet takımının yerine takılması",birim:"adet",birimFiyat:162.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1213,no:"15.465.1115",tanim:"Kavramalı ispanyolet takımının yerine takılması",birim:"adet",birimFiyat:175.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1214,no:"15.465.1116",tanim:"Menteşenin yerine takılması",birim:"adet",birimFiyat:31.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1215,no:"15.465.1117",tanim:"Boy menteşenin yerine takılması",birim:"m",birimFiyat:50.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1216,no:"15.465.1118",tanim:"Ayarlı menteşe (çift) plastik kaplamalının yerine takılması",birim:"adet",birimFiyat:112.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1217,no:"15.465.1201",tanim:"İspanyolet takımının yerine takılması (kol dahil) 100 cm'e kadar, 2 kavramalı",birim:"adet",birimFiyat:437.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1218,no:"15.465.1202",tanim:"İspanyolet takımının yerine takılması (kol dahil) 180 cm'e kadar, 3 kavramalı",birim:"adet",birimFiyat:500.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1219,no:"15.465.1203",tanim:"İspanyolet takımının yerine takılması (kol dahil) 180 cm'den büyük, 4 kavramalı",birim:"adet",birimFiyat:500.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1220,no:"15.465.1204",tanim:"Vasistas ispanyolet takımının yerine takılması (Kol, makas dahil)",birim:"adet",birimFiyat:437.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1221,no:"15.470.1001",tanim:"",birim:"m²",birimFiyat:2335.94,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1222,no:"15.470.1002",tanim:"",birim:"m²",birimFiyat:2381.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1223,no:"15.470.1003",tanim:"",birim:"m²",birimFiyat:2618.13,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1224,no:"15.470.1004",tanim:"",birim:"m²",birimFiyat:2723.13,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1225,no:"15.470.1005",tanim:"",birim:"m²",birimFiyat:2362.19,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1226,no:"15.470.1006",tanim:"",birim:"m²",birimFiyat:2421.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1227,no:"15.470.1007",tanim:"",birim:"m²",birimFiyat:2631.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1228,no:"15.470.1008",tanim:"",birim:"m²",birimFiyat:2736.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1229,no:"15.470.1009",tanim:"",birim:"m²",birimFiyat:1674.38,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1230,no:"15.470.1010",tanim:"",birim:"m²",birimFiyat:1720.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1231,no:"15.470.1011",tanim:"",birim:"m²",birimFiyat:1956.56,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1232,no:"15.470.1012",tanim:"",birim:"m²",birimFiyat:2061.56,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1233,no:"15.470.1013",tanim:"",birim:"m²",birimFiyat:1700.63,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1234,no:"15.470.1014",tanim:"",birim:"m²",birimFiyat:1759.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1235,no:"15.470.1015",tanim:"",birim:"m²",birimFiyat:1969.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1236,no:"15.470.1016",tanim:"",birim:"m²",birimFiyat:2074.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1237,no:"15.470.1201",tanim:"",birim:"m²",birimFiyat:2440.94,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1238,no:"15.470.1202",tanim:"",birim:"m²",birimFiyat:2572.19,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1239,no:"15.470.1203",tanim:"",birim:"m²",birimFiyat:2618.13,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1240,no:"15.470.1204",tanim:"",birim:"m²",birimFiyat:2795.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1241,no:"15.470.1205",tanim:"",birim:"m²",birimFiyat:2703.44,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1242,no:"15.470.1206",tanim:"",birim:"m²",birimFiyat:2506.56,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1243,no:"15.470.1207",tanim:"",birim:"m²",birimFiyat:2585.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1244,no:"15.470.1208",tanim:"",birim:"m²",birimFiyat:2657.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1245,no:"15.470.1209",tanim:"",birim:"m²",birimFiyat:2854.38,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1246,no:"15.470.1210",tanim:"",birim:"m²",birimFiyat:2736.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1247,no:"15.470.1211",tanim:"",birim:"m²",birimFiyat:1779.38,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1248,no:"15.470.1212",tanim:"",birim:"m²",birimFiyat:1910.63,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1249,no:"15.470.1213",tanim:"",birim:"m²",birimFiyat:1956.56,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1250,no:"15.470.1214",tanim:"",birim:"m²",birimFiyat:2133.75,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1251,no:"15.470.1215",tanim:"",birim:"m²",birimFiyat:2041.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1252,no:"15.470.1216",tanim:"",birim:"m²",birimFiyat:1845.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1253,no:"15.470.1217",tanim:"",birim:"m²",birimFiyat:1923.75,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1254,no:"15.470.1218",tanim:"",birim:"m²",birimFiyat:1995.94,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1255,no:"15.470.1219",tanim:"",birim:"m²",birimFiyat:2192.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1256,no:"15.470.1220",tanim:"",birim:"m²",birimFiyat:2074.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1257,no:"15.470.1401",tanim:"",birim:"m²",birimFiyat:2486.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1258,no:"15.470.1402",tanim:"",birim:"m²",birimFiyat:2559.06,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1259,no:"15.470.1403",tanim:"",birim:"m²",birimFiyat:2618.13,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1260,no:"15.470.1404",tanim:"",birim:"m²",birimFiyat:2644.38,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1261,no:"15.470.1405",tanim:"",birim:"m²",birimFiyat:2736.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1262,no:"15.470.1406",tanim:"",birim:"m²",birimFiyat:2821.56,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1263,no:"15.470.1407",tanim:"",birim:"m²",birimFiyat:2605.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1264,no:"15.470.1408",tanim:"",birim:"m²",birimFiyat:2690.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1265,no:"15.470.1409",tanim:"",birim:"m²",birimFiyat:2736.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1266,no:"15.470.1410",tanim:"",birim:"m²",birimFiyat:2762.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1267,no:"15.470.1411",tanim:"",birim:"m²",birimFiyat:2854.38,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1268,no:"15.470.1412",tanim:"",birim:"m²",birimFiyat:2985.63,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1269,no:"15.470.1413",tanim:"",birim:"m²",birimFiyat:1825.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1270,no:"15.470.1414",tanim:"",birim:"m²",birimFiyat:1897.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1271,no:"15.470.1415",tanim:"",birim:"m²",birimFiyat:1956.56,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1272,no:"15.470.1416",tanim:"",birim:"m²",birimFiyat:1982.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1273,no:"15.470.1417",tanim:"",birim:"m²",birimFiyat:2074.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1274,no:"15.470.1418",tanim:"",birim:"m²",birimFiyat:2160.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1275,no:"15.470.1419",tanim:"",birim:"m²",birimFiyat:1943.44,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1276,no:"15.470.1420",tanim:"",birim:"m²",birimFiyat:2028.75,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1277,no:"15.470.1421",tanim:"",birim:"m²",birimFiyat:2074.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1278,no:"15.470.1422",tanim:"",birim:"m²",birimFiyat:2100.94,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1279,no:"15.470.1423",tanim:"",birim:"m²",birimFiyat:2192.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1280,no:"15.470.1424",tanim:"",birim:"m²",birimFiyat:2324.06,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1281,no:"15.475.1001",tanim:"Kadronlu ahşap döşeme yapılması",birim:"m²",birimFiyat:1169.3,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1282,no:"15.475.1002",tanim:"Mevcut kadronlar üzerine ahşap döşeme yapılması",birim:"m²",birimFiyat:953.0,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1283,no:"15.480.1001",tanim:"Beton zemin üzerine kadronlu 15-16 mm kalınlıkta I.sınıf meşe parke kaplama yapılması",birim:"m²",birimFiyat:2865.46,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1284,no:"15.480.1002",tanim:"",birim:"m²",birimFiyat:2615.19,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1285,no:"15.485.1001",tanim:"Lamine parke döşeme kaplaması yapılması (süpürgelik dahil)",birim:"m²",birimFiyat:1712.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1286,no:"15.490.1001",tanim:"Laminat parke döşeme kaplaması yapılması (AC1 Sınıf 21) (süpürgelik dahil)",birim:"m²",birimFiyat:533.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1287,no:"15.490.1002",tanim:"Laminat parke döşeme kaplaması yapılması (AC3 Sınıf 23-31) (süpürgelik dahil)",birim:"m²",birimFiyat:588.51,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1288,no:"15.490.1003",tanim:"Laminat parke döşeme kaplaması yapılması (AC4 Sınıf 32) (süpürgelik dahil)",birim:"m²",birimFiyat:663.06,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1289,no:"15.491.1001",tanim:"",birim:"m²",birimFiyat:636.94,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1290,no:"15.491.1002",tanim:"",birim:"m²",birimFiyat:787.19,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1291,no:"15.491.1003",tanim:"",birim:"m²",birimFiyat:821.56,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1292,no:"15.495.1001",tanim:"Ahşaptan süpürgelik yapılması ve yerine konulması",birim:"m",birimFiyat:141.94,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1293,no:"15.500.1003",tanim:"Düz merdiven küpeştesi yapılması ve yerine konulması",birim:"m",birimFiyat:1133.16,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1294,no:"15.500.1004",tanim:"Eğri merdiven küpeştesi yapılması ve yerine konulması",birim:"m",birimFiyat:1931.9,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1295,no:"15.505.1001",tanim:"Ahşaptan lambri yapılması",birim:"m²",birimFiyat:4524.55,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1296,no:"15.510.1001",tanim:"Ahşaptan masif tablalı iç kapı kasa ve pervazı yapılması yerine konulması",birim:"m²",birimFiyat:2322.39,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1297,no:"15.510.1002",tanim:"Ahşaptan masif tablalı dış kapı kasa ve pervazı yapılması yerine konulması",birim:"m²",birimFiyat:2524.89,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1298,no:"15.510.1101",tanim:"Ahşaptan masif tablalı iç kapı kanadı yapılması ve yerine konulması",birim:"m²",birimFiyat:2158.08,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1299,no:"15.510.1102",tanim:"Ahşaptan masif tablalı dış kapı kanadı yapılması ve yerine konulması",birim:"m²",birimFiyat:2284.83,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1300,no:"15.510.1103",tanim:"",birim:"m²",birimFiyat:3234.69,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1301,no:"15.510.1105",tanim:"Ahşaptan camlı çarpma iç kapı kanadı yapılması ve yerine konulması",birim:"m²",birimFiyat:1894.08,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1302,no:"15.510.9991",tanim:"Mevcut kapılara suni deri ile kapitone kaplama yapılması",birim:"m²",birimFiyat:2794.88,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1303,no:"15.515.1001",tanim:"Ahşaptan kasa ve pervazlı tek satıhlı pencere yapılması ve yerine konulması",birim:"m²",birimFiyat:2123.95,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1304,no:"15.515.1101",tanim:"Ahşaptan iç camekan yapılması ve yerine konulması",birim:"m²",birimFiyat:2003.08,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1305,no:"15.520.1001",tanim:"",birim:"m²",birimFiyat:5696.7,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1306,no:"15.520.1002",tanim:"",birim:"m²",birimFiyat:10518.15,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1307,no:"15.520.1003",tanim:"",birim:"m²",birimFiyat:8254.85,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1308,no:"15.525.1001",tanim:"Ahşap çerçeveli, plastik telden sineklik yapılması ve yerine konulması (takılır-sökülür)",birim:"m²",birimFiyat:1170.11,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1309,no:"15.525.1002",tanim:"Alüminyum çerçeveli, plastik telden sineklik yapılması ve yerine konulması (takılır-sökülür)",birim:"m²",birimFiyat:1152.93,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1310,no:"15.525.1003",tanim:"Pvc çerçeveli, plastik telden sineklik yapılması ve yerine konulması (takılır-sökülür)",birim:"m²",birimFiyat:1123.46,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1311,no:"15.530.1151",tanim:"oranı azaltılmış, kırılma dayanımı artırılmış alçı levhalar ile mevcut duvar üzeri, 60 cm aks",birim:"m²",birimFiyat:1781.79,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1312,no:"15.530.1152",tanim:"oranı azaltılmış, kırılma dayanımı artırılmış alçı levhalar ile mevcut duvar üzeri, 60 cm aks",birim:"m²",birimFiyat:1793.71,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1313,no:"15.530.1251",tanim:"",birim:"m²",birimFiyat:1049.9,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1314,no:"15.530.1252",tanim:"",birim:"m²",birimFiyat:1099.78,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1315,no:"15.530.1253",tanim:"",birim:"m²",birimFiyat:1097.15,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1316,no:"15.530.1254",tanim:"(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1153.59,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1317,no:"15.530.1255",tanim:"",birim:"m²",birimFiyat:1186.14,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1318,no:"15.530.1256",tanim:"",birim:"m²",birimFiyat:1236.01,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1319,no:"15.530.1257",tanim:"",birim:"m²",birimFiyat:1233.39,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1320,no:"15.530.1258",tanim:"(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1289.83,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1321,no:"15.530.1259",tanim:"",birim:"m²",birimFiyat:1090.99,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1322,no:"15.530.1260",tanim:"",birim:"m²",birimFiyat:1140.86,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1323,no:"15.530.1261",tanim:"",birim:"m²",birimFiyat:1138.24,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1324,no:"15.530.1262",tanim:"(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1194.68,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1325,no:"15.530.1263",tanim:"",birim:"m²",birimFiyat:1231.75,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1326,no:"15.530.1264",tanim:"",birim:"m²",birimFiyat:1281.63,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1327,no:"15.530.1265",tanim:"",birim:"m²",birimFiyat:1279.0,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1328,no:"15.530.1266",tanim:"(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1335.44,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1329,no:"15.530.1267",tanim:"",birim:"m²",birimFiyat:1093.78,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1330,no:"15.530.1268",tanim:"",birim:"m²",birimFiyat:1143.65,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1331,no:"15.530.1269",tanim:"",birim:"m²",birimFiyat:1141.03,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1332,no:"15.530.1270",tanim:"(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1197.46,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1333,no:"15.530.1271",tanim:"",birim:"m²",birimFiyat:1240.64,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1334,no:"15.530.1272",tanim:"",birim:"m²",birimFiyat:1290.51,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1335,no:"15.530.1273",tanim:"",birim:"m²",birimFiyat:1287.89,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1336,no:"15.530.1274",tanim:"(her iki yüzünde tek kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1344.33,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1337,no:"15.530.1301",tanim:"",birim:"m²",birimFiyat:1086.65,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1338,no:"15.530.1302",tanim:"",birim:"m²",birimFiyat:1147.03,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1339,no:"15.530.1303",tanim:"",birim:"m²",birimFiyat:1129.96,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1340,no:"15.530.1304",tanim:"(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1190.34,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1341,no:"15.530.1305",tanim:"",birim:"m²",birimFiyat:1222.89,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1342,no:"15.530.1306",tanim:"",birim:"m²",birimFiyat:1283.26,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1343,no:"15.530.1307",tanim:"",birim:"m²",birimFiyat:1266.2,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1344,no:"15.530.1308",tanim:"(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1326.58,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1345,no:"15.530.1309",tanim:"",birim:"m²",birimFiyat:1127.74,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1346,no:"15.530.1310",tanim:"",birim:"m²",birimFiyat:1188.11,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1347,no:"15.530.1311",tanim:"",birim:"m²",birimFiyat:1171.05,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1348,no:"15.530.1312",tanim:"(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1231.43,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1349,no:"15.530.1313",tanim:"",birim:"m²",birimFiyat:1268.5,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1350,no:"15.530.1314",tanim:"",birim:"m²",birimFiyat:1328.88,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1351,no:"15.530.1315",tanim:"",birim:"m²",birimFiyat:1311.81,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1352,no:"15.530.1316",tanim:"(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1372.19,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1353,no:"15.530.1317",tanim:"",birim:"m²",birimFiyat:1130.53,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1354,no:"15.530.1318",tanim:"",birim:"m²",birimFiyat:1190.9,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1355,no:"15.530.1319",tanim:"",birim:"m²",birimFiyat:1173.84,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1356,no:"15.530.1320",tanim:"(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1234.21,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1357,no:"15.530.1321",tanim:"",birim:"m²",birimFiyat:1277.39,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1358,no:"15.530.1322",tanim:"",birim:"m²",birimFiyat:1337.76,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1359,no:"15.530.1323",tanim:"",birim:"m²",birimFiyat:1320.7,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1360,no:"15.530.1324",tanim:"(her iki yüzünde tek kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1381.08,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1361,no:"15.530.1351",tanim:"",birim:"m²",birimFiyat:1306.78,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1362,no:"15.530.1352",tanim:"",birim:"m²",birimFiyat:1406.53,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1363,no:"15.530.1353",tanim:"",birim:"m²",birimFiyat:1401.28,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1364,no:"15.530.1354",tanim:"(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1514.15,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1365,no:"15.530.1355",tanim:"",birim:"m²",birimFiyat:1445.51,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1366,no:"15.530.1356",tanim:"",birim:"m²",birimFiyat:1545.26,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1367,no:"15.530.1357",tanim:"",birim:"m²",birimFiyat:1540.01,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1368,no:"15.530.1358",tanim:"(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1652.89,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1369,no:"15.530.1359",tanim:"",birim:"m²",birimFiyat:1347.86,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1370,no:"15.530.1360",tanim:"",birim:"m²",birimFiyat:1447.61,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1371,no:"15.530.1361",tanim:"",birim:"m²",birimFiyat:1442.36,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1372,no:"15.530.1362",tanim:"(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1555.24,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1373,no:"15.530.1363",tanim:"",birim:"m²",birimFiyat:1491.13,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1374,no:"15.530.1364",tanim:"",birim:"m²",birimFiyat:1590.88,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1375,no:"15.530.1365",tanim:"",birim:"m²",birimFiyat:1585.63,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1376,no:"15.530.1366",tanim:"(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1698.5,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1377,no:"15.530.1367",tanim:"",birim:"m²",birimFiyat:1350.65,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1378,no:"15.530.1368",tanim:"",birim:"m²",birimFiyat:1450.4,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1379,no:"15.530.1369",tanim:"",birim:"m²",birimFiyat:1445.15,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1380,no:"15.530.1370",tanim:"(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1558.03,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1381,no:"15.530.1371",tanim:"",birim:"m²",birimFiyat:1500.01,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1382,no:"15.530.1372",tanim:"",birim:"m²",birimFiyat:1599.76,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1383,no:"15.530.1373",tanim:"",birim:"m²",birimFiyat:1594.51,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1384,no:"15.530.1374",tanim:"(her iki yüzünde çift kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1707.39,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1385,no:"15.530.1401",tanim:"",birim:"m²",birimFiyat:1380.28,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1386,no:"15.530.1402",tanim:"",birim:"m²",birimFiyat:1501.03,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1387,no:"15.530.1403",tanim:"",birim:"m²",birimFiyat:1466.9,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1388,no:"15.530.1404",tanim:"(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1587.65,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1389,no:"15.530.1405",tanim:"",birim:"m²",birimFiyat:1519.01,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1390,no:"15.530.1406",tanim:"",birim:"m²",birimFiyat:1639.76,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1391,no:"15.530.1407",tanim:"",birim:"m²",birimFiyat:1605.64,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1392,no:"15.530.1408",tanim:"(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1726.39,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1393,no:"15.530.1409",tanim:"",birim:"m²",birimFiyat:1421.36,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1394,no:"15.530.1410",tanim:"",birim:"m²",birimFiyat:1542.11,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1395,no:"15.530.1411",tanim:"",birim:"m²",birimFiyat:1507.99,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1396,no:"15.530.1412",tanim:"(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1628.74,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1397,no:"15.530.1413",tanim:"",birim:"m²",birimFiyat:1564.63,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1398,no:"15.530.1414",tanim:"",birim:"m²",birimFiyat:1685.38,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1399,no:"15.530.1415",tanim:"",birim:"m²",birimFiyat:1651.25,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1400,no:"15.530.1416",tanim:"(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1772.0,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1401,no:"15.530.1417",tanim:"",birim:"m²",birimFiyat:1424.15,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1402,no:"15.530.1418",tanim:"",birim:"m²",birimFiyat:1544.9,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1403,no:"15.530.1419",tanim:"",birim:"m²",birimFiyat:1510.78,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1404,no:"15.530.1420",tanim:"(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1631.53,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1405,no:"15.530.1421",tanim:"",birim:"m²",birimFiyat:1573.51,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1406,no:"15.530.1422",tanim:"",birim:"m²",birimFiyat:1694.26,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1407,no:"15.530.1423",tanim:"",birim:"m²",birimFiyat:1660.14,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1408,no:"15.530.1424",tanim:"(her iki yüzünde çift kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1780.89,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1409,no:"15.530.1451",tanim:"",birim:"m²",birimFiyat:1564.9,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1410,no:"15.530.1452",tanim:"",birim:"m²",birimFiyat:1714.53,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1411,no:"15.530.1453",tanim:"",birim:"m²",birimFiyat:1706.65,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1412,no:"15.530.1454",tanim:"(her iki yüzünde üç kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1875.96,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1413,no:"15.530.1455",tanim:"",birim:"m²",birimFiyat:1605.99,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1414,no:"15.530.1456",tanim:"",birim:"m²",birimFiyat:1755.61,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1415,no:"15.530.1457",tanim:"",birim:"m²",birimFiyat:1747.74,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1416,no:"15.530.1458",tanim:"(her iki yüzünde üç kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1917.05,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1417,no:"15.530.1459",tanim:"",birim:"m²",birimFiyat:1608.78,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1418,no:"15.530.1460",tanim:"",birim:"m²",birimFiyat:1758.4,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1419,no:"15.530.1461",tanim:"",birim:"m²",birimFiyat:1750.53,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1420,no:"15.530.1462",tanim:"(her iki yüzünde üç kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1919.84,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1421,no:"15.530.1501",tanim:"",birim:"m²",birimFiyat:1733.36,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1422,no:"15.530.1502",tanim:"",birim:"m²",birimFiyat:1833.11,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1423,no:"15.530.1503",tanim:"",birim:"m²",birimFiyat:1827.86,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1424,no:"15.530.1504",tanim:"(her iki yüzünde iki kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:1940.74,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1425,no:"15.530.1505",tanim:"",birim:"m²",birimFiyat:1806.86,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1426,no:"15.530.1506",tanim:"",birim:"m²",birimFiyat:1927.61,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1427,no:"15.530.1507",tanim:"",birim:"m²",birimFiyat:1893.49,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1428,no:"15.530.1508",tanim:"(her iki yüzünde iki kat 15 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:2014.24,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1429,no:"15.530.1509",tanim:"",birim:"m²",birimFiyat:1815.51,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1430,no:"15.530.1510",tanim:"",birim:"m²",birimFiyat:1915.26,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1431,no:"15.530.1511",tanim:"",birim:"m²",birimFiyat:1910.01,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1432,no:"15.530.1512",tanim:"(her iki yüzünde iki kat 12,5 mm su emme oranı azaltılmış, yangına dayanımı artırılmış alçı",birim:"m²",birimFiyat:2022.89,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1433,no:"15.530.1551",tanim:"(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm standart alçı levha",birim:"m²",birimFiyat:1976.24,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1434,no:"15.530.1552",tanim:"(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm su emme oranı",birim:"m²",birimFiyat:2078.36,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1435,no:"15.530.1553",tanim:"(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm yangına dayanımı",birim:"m²",birimFiyat:2072.99,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1436,no:"15.530.1554",tanim:"(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm su emme oranı",birim:"m²",birimFiyat:2188.55,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1437,no:"15.530.1555",tanim:"(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 15 mm standart alçı levha ile)",birim:"m²",birimFiyat:2051.49,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1438,no:"15.530.1556",tanim:"(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 15 mm su emme oranı",birim:"m²",birimFiyat:2175.11,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1439,no:"15.530.1557",tanim:"(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 15 mm yangına dayanımı",birim:"m²",birimFiyat:2140.18,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1440,no:"15.530.1558",tanim:"(Duvar C 50 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 15 mm su emme oranı",birim:"m²",birimFiyat:2263.8,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1441,no:"15.530.1559",tanim:"(Duvar C 75 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm standart alçı levha",birim:"m²",birimFiyat:2058.39,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1442,no:"15.530.1560",tanim:"(Duvar C 75 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm su emme oranı",birim:"m²",birimFiyat:2160.51,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1443,no:"15.530.1561",tanim:"(Duvar C 75 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm yangına dayanımı",birim:"m²",birimFiyat:2155.14,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1444,no:"15.530.1562",tanim:"(Duvar C 75 profil - 60 cm aks aralığı) (her iki yüzünde iki kat 12,5 mm su emme oranı",birim:"m²",birimFiyat:2270.7,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1445,no:"15.530.1701",tanim:"",birim:"m²",birimFiyat:903.24,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1446,no:"15.530.1702",tanim:"",birim:"m²",birimFiyat:928.18,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1447,no:"15.530.1703",tanim:"",birim:"m²",birimFiyat:926.86,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1448,no:"15.530.1704",tanim:"",birim:"m²",birimFiyat:955.09,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1449,no:"15.530.1727",tanim:"",birim:"m²",birimFiyat:694.44,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1450,no:"15.530.1728",tanim:"",birim:"m²",birimFiyat:693.13,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1451,no:"15.530.1729",tanim:"",birim:"m²",birimFiyat:721.35,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1452,no:"15.530.1731",tanim:"",birim:"m²",birimFiyat:718.06,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1453,no:"15.530.1732",tanim:"",birim:"m²",birimFiyat:709.54,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1454,no:"15.530.1733",tanim:"",birim:"m²",birimFiyat:739.73,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1455,no:"15.530.1752",tanim:"",birim:"m²",birimFiyat:915.44,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1456,no:"15.530.1753",tanim:"",birim:"m²",birimFiyat:912.81,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1457,no:"15.530.1754",tanim:"aralığı) (12,5 mm çift kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha",birim:"m²",birimFiyat:969.25,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1458,no:"15.530.1776",tanim:"",birim:"m²",birimFiyat:848.36,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1459,no:"15.530.1777",tanim:"",birim:"m²",birimFiyat:873.3,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1460,no:"15.530.1778",tanim:"",birim:"m²",birimFiyat:871.99,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1461,no:"15.530.1779",tanim:"aralığı) (12,5 mm tek kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha",birim:"m²",birimFiyat:900.21,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1462,no:"15.530.1780",tanim:"",birim:"m²",birimFiyat:1017.1,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1463,no:"15.530.1781",tanim:"",birim:"m²",birimFiyat:1042.04,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1464,no:"15.530.1782",tanim:"",birim:"m²",birimFiyat:1040.73,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1465,no:"15.530.1783",tanim:"aralığı) (12,5 mm tek kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha",birim:"m²",birimFiyat:1068.95,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1466,no:"15.530.1784",tanim:"",birim:"m²",birimFiyat:889.45,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1467,no:"15.530.1785",tanim:"",birim:"m²",birimFiyat:914.39,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1468,no:"15.530.1786",tanim:"",birim:"m²",birimFiyat:913.08,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1469,no:"15.530.1787",tanim:"aralığı) (12,5 mm tek kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha",birim:"m²",birimFiyat:941.3,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1470,no:"15.530.1788",tanim:"",birim:"m²",birimFiyat:1062.71,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1471,no:"15.530.1789",tanim:"",birim:"m²",birimFiyat:1087.65,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1472,no:"15.530.1790",tanim:"",birim:"m²",birimFiyat:1086.34,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1473,no:"15.530.1791",tanim:"aralığı) (12,5 mm tek kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha",birim:"m²",birimFiyat:1114.56,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1474,no:"15.530.1792",tanim:"",birim:"m²",birimFiyat:866.74,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1475,no:"15.530.1793",tanim:"",birim:"m²",birimFiyat:929.41,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1476,no:"15.530.1794",tanim:"",birim:"m²",birimFiyat:888.4,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1477,no:"15.530.1795",tanim:"",birim:"m²",birimFiyat:918.59,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1478,no:"15.530.1796",tanim:"",birim:"m²",birimFiyat:1035.48,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1479,no:"15.530.1797",tanim:"",birim:"m²",birimFiyat:1065.66,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1480,no:"15.530.1798",tanim:"",birim:"m²",birimFiyat:1057.14,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1481,no:"15.530.1799",tanim:"",birim:"m²",birimFiyat:1087.33,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1482,no:"15.530.1800",tanim:"",birim:"m²",birimFiyat:907.83,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1483,no:"15.530.1801",tanim:"",birim:"m²",birimFiyat:938.01,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1484,no:"15.530.1802",tanim:"",birim:"m²",birimFiyat:929.49,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1485,no:"15.530.1803",tanim:"",birim:"m²",birimFiyat:959.68,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1486,no:"15.530.1804",tanim:"",birim:"m²",birimFiyat:1081.09,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1487,no:"15.530.1805",tanim:"",birim:"m²",birimFiyat:1111.28,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1488,no:"15.530.1806",tanim:"",birim:"m²",birimFiyat:1102.75,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1489,no:"15.530.1807",tanim:"",birim:"m²",birimFiyat:1132.94,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1490,no:"15.530.1826",tanim:"",birim:"m²",birimFiyat:1044.43,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1491,no:"15.530.1827",tanim:"",birim:"m²",birimFiyat:1094.3,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1492,no:"15.530.1828",tanim:"",birim:"m²",birimFiyat:1091.68,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1493,no:"15.530.1829",tanim:"aralığı) (12,5 mm çift kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha",birim:"m²",birimFiyat:1148.11,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1494,no:"15.530.1830",tanim:"",birim:"m²",birimFiyat:1079.41,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1495,no:"15.530.1831",tanim:"",birim:"m²",birimFiyat:1129.29,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1496,no:"15.530.1832",tanim:"",birim:"m²",birimFiyat:1126.66,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1497,no:"15.530.1833",tanim:"aralığı) (12,5 mm çift kat su emme oranı azaltılmış, yangına dayanımı artırılmış alçı levha",birim:"m²",birimFiyat:1183.1,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1498,no:"15.535.1001",tanim:"(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 serisi) oturmalı sistem asma",birim:"m²",birimFiyat:1098.74,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1499,no:"15.535.1002",tanim:"boyalı(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) oturmalı sistem",birim:"m²",birimFiyat:1124.99,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1500,no:"15.535.1003",tanim:"(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000",birim:"m²",birimFiyat:1216.86,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1501,no:"15.535.1004",tanim:"(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 serisi) oturmalı sistem asma",birim:"m²",birimFiyat:1517.86,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1502,no:"15.535.1005",tanim:"(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 SERISI) oturmalı sistem asma",birim:"m²",birimFiyat:1544.11,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1503,no:"15.535.1006",tanim:"(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) oturmalı sistem asma",birim:"m²",birimFiyat:1524.43,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1504,no:"15.535.1007",tanim:"(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) oturmalı sistem asma",birim:"m²",birimFiyat:1544.11,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1505,no:"15.535.1008",tanim:"(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000",birim:"m²",birimFiyat:1544.11,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1506,no:"15.535.1009",tanim:"(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000",birim:"m²",birimFiyat:1570.36,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1507,no:"15.535.1010",tanim:"(polyester esaslı) deliksiz sıcak daldırma galvanize sac plakadan oturmalı sistem asma tavan",birim:"m²",birimFiyat:908.43,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1508,no:"15.535.1011",tanim:"(polyester esaslı) delikli sıcak daldırma galvanize sac plakadan oturmalı sistem asma tavan",birim:"m²",birimFiyat:914.99,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1509,no:"15.535.1012",tanim:"(polyester esaslı) arka yüzü akustik kumaş kaplı delikli sıcak daldrıma galvanize sac",birim:"m²",birimFiyat:954.36,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1510,no:"15.535.1013",tanim:"(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 serisi) sarkmalı sistem asma",birim:"m²",birimFiyat:1105.4,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1511,no:"15.535.1014",tanim:"(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) sarkmalı sistem asma",birim:"m²",birimFiyat:1131.65,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1512,no:"15.535.1015",tanim:"(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000",birim:"m²",birimFiyat:1197.28,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1513,no:"15.535.1016",tanim:"(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 serisi)",birim:"m²",birimFiyat:1524.53,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1514,no:"15.535.1017",tanim:"(polyester esaslı) deliksiz alüminyum plakadan (EN AW 3000 SERISI)",birim:"m²",birimFiyat:1531.09,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1515,no:"15.535.1018",tanim:"(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) sarkmalı sistem asma",birim:"m²",birimFiyat:1524.53,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1516,no:"15.535.1019",tanim:"(polyester esaslı) delikli alüminyum plakadan (EN AW 3000 serisi) sarkmalı sistem asma",birim:"m²",birimFiyat:1550.78,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1517,no:"15.535.1020",tanim:"(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000",birim:"m²",birimFiyat:1583.59,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1518,no:"15.535.1021",tanim:"(polyester esaslı) arka yüzü akustik kumaş kaplı delikli alüminyum plakadan (EN AW 3000",birim:"m²",birimFiyat:1583.59,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1519,no:"15.535.1022",tanim:"(polyester esaslı) deliksiz sıcak daldırma galvanize sac plakadan sarkmalı sistem asma tavan",birim:"m²",birimFiyat:869.15,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1520,no:"15.535.1023",tanim:"(polyester esaslı) delikli sıcak daldırma galvanize sac plakadan sarkmalı sistem asma tavan",birim:"m²",birimFiyat:908.53,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1521,no:"15.535.1024",tanim:"(polyester esaslı) arka yüzü akustik kumaş kaplı delikli sıcak daldırma galvanize sac",birim:"m²",birimFiyat:954.46,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1522,no:"15.535.1036",tanim:"",birim:"m²",birimFiyat:1283.6,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1523,no:"15.535.1037",tanim:"",birim:"m²",birimFiyat:1435.48,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1524,no:"15.535.1038",tanim:"",birim:"m²",birimFiyat:1197.98,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1525,no:"15.535.1039",tanim:"",birim:"m²",birimFiyat:1344.79,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1526,no:"15.540.1014",tanim:"Ahşap yüzeylerin verniklenmesi",birim:"m²",birimFiyat:292.98,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1527,no:"15.540.1015",tanim:"Ahşap yüzeylerin vernikli renkli ahşap koruyucu ile verniklenmesi",birim:"m²",birimFiyat:300.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1528,no:"15.540.1016",tanim:"Ahşap yüzeylerin renkli ahşap koruyucu ile korunması",birim:"m²",birimFiyat:273.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1529,no:"15.540.1017",tanim:"Her cins ahşap parkenin cilalanması",birim:"m²",birimFiyat:425.94,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1530,no:"15.540.1021",tanim:"Ahşap yüzeylere bir kat sentetik boya yapılması",birim:"m²",birimFiyat:237.35,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1531,no:"15.540.1022",tanim:"Ahşap yüzeylere bir kat sentetik esaslı mat boya yapılması",birim:"m²",birimFiyat:231.6,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1532,no:"15.540.1023",tanim:"Ahşap yüzeylere iki kat sentetik esaslı parlak boya yapılması",birim:"m²",birimFiyat:309.54,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1533,no:"15.540.1024",tanim:"Ahşap yüzeylere iki kat sentetik esaslı mat boya yapılması",birim:"m²",birimFiyat:298.04,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1534,no:"15.540.1025",tanim:"",birim:"m²",birimFiyat:479.21,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1535,no:"15.540.1111",tanim:"Demir yüzeylere korozyona karşı iki kat antipas boya yapılması",birim:"m²",birimFiyat:270.38,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1536,no:"15.540.1121",tanim:"Demir yüzeylere iki kat antipas, iki kat sentetik esaslı parlak boya yapılması",birim:"m²",birimFiyat:375.38,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1537,no:"15.540.1122",tanim:"Demir yüzeylere iki kat antipas, iki kat sentetik esaslı mat boya yapılması",birim:"m²",birimFiyat:330.55,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1538,no:"15.540.1123",tanim:"Demir yüzeylere iki kat solvent bazlı epoksi boya yapılması",birim:"m²",birimFiyat:408.36,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1539,no:"15.540.1421",tanim:"Beton, sıva ve benzeri yüzeylere 1,5 mm kalınlıkta akrilik esaslı renkli kaplama yapılması",birim:"m²",birimFiyat:567.19,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1540,no:"15.540.1422",tanim:"Beton, sıva ve benzeri yüzeylere 2 mm kalınlıkta akrilik esaslı renkli kaplama yapılması",birim:"m²",birimFiyat:629.69,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1541,no:"15.540.1423",tanim:"Beton, sıva ve benzeri yüzeylere 3 mm kalınlıkta akrilik esaslı renkli kaplama yapılması",birim:"m²",birimFiyat:731.25,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1542,no:"15.540.1424",tanim:"",birim:"m²",birimFiyat:591.94,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1543,no:"15.540.1425",tanim:"",birim:"m²",birimFiyat:665.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1544,no:"15.540.1426",tanim:"",birim:"m²",birimFiyat:776.25,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1545,no:"15.540.1427",tanim:"Beton, sıva ve benzeri yüzeylere 1,5 mm kalınlıkta çimento esaslı kaplama yapılması",birim:"m²",birimFiyat:465.38,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1546,no:"15.540.1428",tanim:"Beton, sıva ve benzeri yüzeylere 2 mm kalınlıkta çimento esaslı kaplama yapılması",birim:"m²",birimFiyat:469.15,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1547,no:"15.540.1429",tanim:"Beton, sıva ve benzeri yüzeylere 3 mm kalınlıkta çimento esaslı kaplama yapılması",birim:"m²",birimFiyat:543.74,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1548,no:"15.540.1501",tanim:"Brüt beton yüzeylerin alçı veya sıvaya hazırlanması (iç cephe)",birim:"m²",birimFiyat:134.06,idare:"Çevre ve Şehircilik",kategori:"Beton İşleri",kaynak:"2026"},
  {id:1549,no:"15.540.1502",tanim:"Lekeli ve isli duvar yüzeylerin boya işlemine hazır hale getirilmesi (iç cephe)",birim:"m²",birimFiyat:276.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1550,no:"15.540.1503",tanim:"Eski boyalı yüzeylere üç kat beyaz kireç badana yapılması (iç cephe)",birim:"m²",birimFiyat:259.98,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1551,no:"15.540.1504",tanim:"Eski boyalı yüzeylere üç kat renkli kireç badana yapılması (iç cephe)",birim:"m²",birimFiyat:260.73,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1552,no:"15.540.1505",tanim:"",birim:"m²",birimFiyat:217.83,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1553,no:"15.540.1506",tanim:"",birim:"m²",birimFiyat:233.83,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1554,no:"15.540.1507",tanim:"",birim:"m²",birimFiyat:216.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1555,no:"15.540.1508",tanim:"",birim:"m²",birimFiyat:222.51,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1556,no:"15.540.1509",tanim:"",birim:"m²",birimFiyat:228.51,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1557,no:"15.540.1510",tanim:"",birim:"m²",birimFiyat:264.71,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1558,no:"15.540.1511",tanim:"",birim:"m²",birimFiyat:268.76,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1559,no:"15.540.1512",tanim:"",birim:"m²",birimFiyat:310.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1560,no:"15.540.1513",tanim:"",birim:"m²",birimFiyat:323.2,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1561,no:"15.540.1514",tanim:"",birim:"m²",birimFiyat:310.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1562,no:"15.540.1515",tanim:"Yeni sıva yüzeylere üç kat beyaz kireç badana yapılması (iç cephe)",birim:"m²",birimFiyat:258.99,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1563,no:"15.540.1516",tanim:"Yeni sıva yüzeylere üç kat renkli kireç badana yapılması (iç cephe)",birim:"m²",birimFiyat:259.74,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1564,no:"15.540.1517",tanim:"",birim:"m²",birimFiyat:377.59,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1565,no:"15.540.1518",tanim:"",birim:"m²",birimFiyat:380.78,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1566,no:"15.540.1519",tanim:"",birim:"m²",birimFiyat:373.43,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1567,no:"15.540.1520",tanim:"",birim:"m²",birimFiyat:216.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1568,no:"15.540.1521",tanim:"",birim:"m²",birimFiyat:223.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1569,no:"15.540.1522",tanim:"",birim:"m²",birimFiyat:215.04,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1570,no:"15.540.1523",tanim:"",birim:"m²",birimFiyat:182.31,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1571,no:"15.540.1524",tanim:"",birim:"m²",birimFiyat:185.5,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1572,no:"15.540.1525",tanim:"",birim:"m²",birimFiyat:178.15,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1573,no:"15.540.1526",tanim:"",birim:"m²",birimFiyat:389.71,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1574,no:"15.540.1527",tanim:"",birim:"m²",birimFiyat:224.63,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1575,no:"15.540.1528",tanim:"",birim:"m²",birimFiyat:187.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1576,no:"15.540.1529",tanim:"",birim:"m²",birimFiyat:388.28,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1577,no:"15.540.1530",tanim:"",birim:"m²",birimFiyat:228.13,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1578,no:"15.540.1531",tanim:"",birim:"m²",birimFiyat:193.0,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1579,no:"15.540.1532",tanim:"Antibakteriyel İç Cephe Boyası (Örtücülük Sınıf:2, YOD:Sınıf2, Parlaklık:G3) yapılması (iç",birim:"m²",birimFiyat:376.6,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1580,no:"15.540.1533",tanim:"Antibakteriyel İç Cephe Boyası (Örtücülük Sınıf:2, YOD:Sınıf1, Parlaklık:G2) yapılması (iç",birim:"m²",birimFiyat:384.88,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1581,no:"15.540.1534",tanim:"",birim:"m²",birimFiyat:369.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1582,no:"15.540.1535",tanim:"",birim:"m²",birimFiyat:305.58,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1583,no:"15.540.1536",tanim:"",birim:"m²",birimFiyat:309.63,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1584,no:"15.540.1537",tanim:"",birim:"m²",birimFiyat:296.13,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1585,no:"15.540.1538",tanim:"Antibakteriyel İç Cephe Boyası (Örtücülük Sınıf:2, YOD:Sınıf2, Parlaklık:G3) yapılması (iç",birim:"m²",birimFiyat:240.09,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1586,no:"15.540.1539",tanim:"Antibakteriyel İç Cephe Boyası (Örtücülük Sınıf:2, YOD:Sınıf1, Parlaklık:G2) yapılması (iç",birim:"m²",birimFiyat:247.44,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1587,no:"15.540.1540",tanim:"",birim:"m²",birimFiyat:233.94,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1588,no:"15.540.1541",tanim:"",birim:"m²",birimFiyat:369.09,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1589,no:"15.540.1542",tanim:"",birim:"m²",birimFiyat:355.89,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1590,no:"15.540.1543",tanim:"Yeni sıva yüzeylere astar uygulanarak iki kat sentetik esaslı parlak boya yapılması (iç cephe)",birim:"m²",birimFiyat:371.19,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1591,no:"15.540.1544",tanim:"Yeni sıva yüzeylere astar uygulanarak iki kat sentetik esaslı mat boya yapılması (iç cephe)",birim:"m²",birimFiyat:357.99,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1592,no:"15.540.1545",tanim:"",birim:"m²",birimFiyat:351.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1593,no:"15.540.1546",tanim:"",birim:"m²",birimFiyat:338.61,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1594,no:"15.540.1547",tanim:"",birim:"m²",birimFiyat:225.08,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1595,no:"15.540.1548",tanim:"",birim:"m²",birimFiyat:253.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1596,no:"15.540.1549",tanim:"",birim:"m²",birimFiyat:224.69,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1597,no:"15.540.1601",tanim:"",birim:"m²",birimFiyat:359.83,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1598,no:"15.540.1602",tanim:"",birim:"m²",birimFiyat:386.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1599,no:"15.540.1603",tanim:"",birim:"m²",birimFiyat:389.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1600,no:"15.540.1604",tanim:"",birim:"m²",birimFiyat:391.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1601,no:"15.540.1605",tanim:"",birim:"m²",birimFiyat:391.03,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1602,no:"15.540.1606",tanim:"",birim:"m²",birimFiyat:386.03,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1603,no:"15.540.1607",tanim:"",birim:"m²",birimFiyat:502.59,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1604,no:"15.540.1608",tanim:"",birim:"m²",birimFiyat:386.03,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1605,no:"15.540.1609",tanim:"",birim:"m²",birimFiyat:419.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1606,no:"15.540.1610",tanim:"",birim:"m²",birimFiyat:383.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1607,no:"15.540.1611",tanim:"",birim:"m²",birimFiyat:380.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1608,no:"15.540.1612",tanim:"",birim:"m²",birimFiyat:411.89,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1609,no:"15.540.1613",tanim:"",birim:"m²",birimFiyat:414.89,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1610,no:"15.540.1614",tanim:"",birim:"m²",birimFiyat:416.89,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1611,no:"15.540.1615",tanim:"",birim:"m²",birimFiyat:383.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1612,no:"15.540.1616",tanim:"",birim:"m²",birimFiyat:380.81,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1613,no:"15.540.1617",tanim:"",birim:"m²",birimFiyat:386.03,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1614,no:"15.540.1618",tanim:"",birim:"m²",birimFiyat:402.28,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1615,no:"15.540.1619",tanim:"",birim:"m²",birimFiyat:386.03,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1616,no:"15.540.1620",tanim:"",birim:"m²",birimFiyat:402.28,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1617,no:"15.540.1621",tanim:"",birim:"m²",birimFiyat:386.03,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1618,no:"15.540.1622",tanim:"",birim:"m²",birimFiyat:402.28,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1619,no:"15.540.1623",tanim:"",birim:"m²",birimFiyat:503.39,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1620,no:"15.540.1624",tanim:"",birim:"m²",birimFiyat:514.08,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1621,no:"15.540.1625",tanim:"",birim:"m²",birimFiyat:362.6,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1622,no:"15.550.1001",tanim:"Kare ve dikdörtgen profillerle pencere ve kapı yapılması ve yerine konulması",birim:"kg",birimFiyat:183.55,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1623,no:"15.550.1002",tanim:"",birim:"kg",birimFiyat:218.54,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1624,no:"15.550.1003",tanim:"",birim:"kg",birimFiyat:216.78,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1625,no:"15.550.1004",tanim:"1,50 mm kalınlığında düz siyah sacdan bükme kapı kasası yapılması ve yerine konulması",birim:"kg",birimFiyat:218.54,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1626,no:"15.550.1005",tanim:"2,00 mm kalınlığında düz siyah sacdan bükme kapı kasası yapılması ve yerine konulması",birim:"kg",birimFiyat:216.78,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1627,no:"15.550.1201",tanim:"",birim:"kg",birimFiyat:166.28,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1628,no:"15.550.1202",tanim:"Lama ve profil demirlerden çeşitli demir işleri yapılması ve yerine konulması",birim:"kg",birimFiyat:175.8,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1629,no:"15.550.1203",tanim:"Demir borudan kaynakla korkuluk yapılması, yerine konulması",birim:"kg",birimFiyat:170.51,idare:"Çevre ve Şehircilik",kategori:"Demir İşleri",kaynak:"2026"},
  {id:1630,no:"15.550.1204",tanim:"",birim:"kg",birimFiyat:155.39,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1631,no:"15.555.1001",tanim:"daldırma galvaniz üzeri elektrostatik polyester toz boyalı panel teller ile çit yapılması (Direk",birim:"m",birimFiyat:693.2,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1632,no:"15.555.1002",tanim:"daldırma galvaniz üzeri elektrostatik polyester toz boyalı panel teller ile çit yapılması (Direk",birim:"m",birimFiyat:810.95,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1633,no:"15.555.1003",tanim:"daldırma galvaniz üzeri elektrostatik polyester toz boyalı panel teller ile çit yapılması (Direk",birim:"m",birimFiyat:917.2,idare:"Çevre ve Şehircilik",kategori:"İnce İnşaat",kaynak:"2026"},
  {id:1634,no:"15.560.1001",tanim:"Fonttan; ızgara, kapak, garguy yapılması ve yerine konulması",birim:"kg",birimFiyat:127.5,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1635,no:"15.560.1002",tanim:"",birim:"adet",birimFiyat:5270.0,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1636,no:"15.560.1003",tanim:"",birim:"adet",birimFiyat:4457.5,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1637,no:"15.560.1004",tanim:"",birim:"adet",birimFiyat:4957.5,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1638,no:"15.560.2001",tanim:"",birim:"m³",birimFiyat:153.75,idare:"Çevre ve Şehircilik",kategori:"Diğer",kaynak:"2026"},
  {id:1639,no:"19.100.1001",tanim:"Ekskavatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:2183.46,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1640,no:"19.100.1002",tanim:"Ekskavatör Beko'nun 1 saatlik ücreti",birim:"Sa",birimFiyat:2571.91,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1641,no:"19.100.1003",tanim:"Ekskavatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:2645.46,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1642,no:"19.100.1004",tanim:"Ekskavatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:2867.81,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1643,no:"19.100.1005",tanim:"Ekskavatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:3374.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1644,no:"19.100.1006",tanim:"Ekskavatör (paletli) (210-259 HP) (maksimum 2,5 m³) 1 saatlik ücreti",birim:"Sa",birimFiyat:2971.46,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1645,no:"19.100.1007",tanim:"Ekskavatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:3836.26,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1646,no:"19.100.1008",tanim:"Ekskavatör (paletli) (260-299 HP) (maksimum 2,5 m³) 1 saatlik ücreti",birim:"Sa",birimFiyat:3421.44,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1647,no:"19.100.1009",tanim:"Ekskavatör (paletli) (300-329 HP) (maksimum 3,5 m³) 1 saatlik ücreti",birim:"Sa",birimFiyat:3984.58,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1648,no:"19.100.1010",tanim:"Traktör Skreyper'in 1 saatlik ücreti",birim:"Sa",birimFiyat:2198.58,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1649,no:"19.100.1011",tanim:"Traktör Ripper'in 1 saatlik ücreti",birim:"Sa",birimFiyat:3389.66,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1650,no:"19.100.1012",tanim:"Motor Greyder'in 1 saatlik ücreti",birim:"Sa",birimFiyat:1830.41,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1651,no:"19.100.1013",tanim:"Greyder (190-209 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:2622.44,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1652,no:"19.100.1014",tanim:"Greyder (210-230 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:2933.93,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1653,no:"19.100.1015",tanim:"Lastik tekerlekli traktör Skreyper'in 1 saatlik ücreti",birim:"Sa",birimFiyat:4947.48,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1654,no:"19.100.1016",tanim:"Traktör Buldozer'in (70 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:1789.23,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1655,no:"19.100.1017",tanim:"Traktör Buldozer'in (100 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:2049.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1656,no:"19.100.1018",tanim:"Traktör Buldozer'in (160 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:2586.03,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1657,no:"19.100.1019",tanim:"Traktör Buldozer'in 1 saatlik ücreti",birim:"Sa",birimFiyat:3141.71,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1658,no:"19.100.1020",tanim:"Traktör Buldozer'in (285 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:4868.73,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1659,no:"19.100.1021",tanim:"Traktör Buldozer'in (345 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:5439.9,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1660,no:"19.100.1022",tanim:"Şahmerdan'ın (50 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:2540.26,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1661,no:"19.100.1023",tanim:"Kompresör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:2508.08,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1662,no:"19.100.1024",tanim:"Vantilatasyon için Kompresör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:1461.78,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1663,no:"19.100.1025",tanim:"Kompresör'ün 1 saatlik ücreti (250 HP)",birim:"Sa",birimFiyat:1942.71,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1664,no:"19.100.1026",tanim:"Enjeksiyon Makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:1471.78,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1665,no:"19.100.1027",tanim:"Kazıcı yükleyici (100 HP) (maksimum 2,5 m³) 1 saatlik ücreti",birim:"Sa",birimFiyat:1656.24,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1666,no:"19.100.1028",tanim:"Lastik tekerlekli yükleyicinin 1 saatlik ücreti",birim:"Sa",birimFiyat:1686.26,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1667,no:"19.100.1029",tanim:"Yükleyici (lastik tekerlekli) (100 HP) (maksimum 2 m³) 1 saatlik ücreti",birim:"Sa",birimFiyat:1510.89,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1668,no:"19.100.1030",tanim:"Paletli Yükleyicinin 1 saatlik ücreti",birim:"Sa",birimFiyat:2037.39,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1669,no:"19.100.1031",tanim:"Betoniyer'in 1 saatlik ücreti",birim:"Sa",birimFiyat:1012.75,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1670,no:"19.100.1032",tanim:"Mozayik silme makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:528.95,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1671,no:"19.100.1033",tanim:"Beton vibratörü'nün 1 saatlik ücreti",birim:"Sa",birimFiyat:458.72,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1672,no:"19.100.1034",tanim:"Konkasör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:4659.08,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1673,no:"19.100.1035",tanim:"Elek makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:1334.95,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1674,no:"19.100.1036",tanim:"Eleme makinasının 1 saatlik ücreti",birim:"Sa",birimFiyat:1254.75,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1675,no:"19.100.1037",tanim:"Motopomp'un (5 Ps.) 1 saatlik ücreti",birim:"Sa",birimFiyat:305.51,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1676,no:"19.100.1038",tanim:"Motopompun 1 saatlik ücreti",birim:"Sa",birimFiyat:364.62,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1677,no:"19.100.1039",tanim:"Motopomp'un (20 Ps.) 1 saatlik ücreti",birim:"Sa",birimFiyat:444.43,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1678,no:"19.100.1040",tanim:"Motopomp'un (30 Ps.) 1 saatlik ücreti",birim:"Sa",birimFiyat:735.16,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1679,no:"19.100.1041",tanim:"Motopomp'un (45 Ps.) 1 saatlik ücreti",birim:"Sa",birimFiyat:874.78,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1680,no:"19.100.1042",tanim:"Motopomp'un (60 Ps.) 1 saatlik ücreti",birim:"Sa",birimFiyat:1004.89,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1681,no:"19.100.1043",tanim:"Çekilir tip beton pompası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:1697.1,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1682,no:"19.100.1044",tanim:"Arazöz'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:882.41,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1683,no:"19.100.1045",tanim:"Arazöz Pick-Up'ın 1 saatlik ücreti",birim:"Sa",birimFiyat:845.99,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1684,no:"19.100.1046",tanim:"Her cins (Titreşimli Darbeli) motorlu kompaktör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:978.34,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1685,no:"19.100.1047",tanim:"Titreşimli Silindir'in 1 saatlik ücreti",birim:"Sa",birimFiyat:1531.16,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1686,no:"19.100.1048",tanim:"Titreşimli Silindir'in 1 saatlik ücreti",birim:"Sa",birimFiyat:1597.17,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1687,no:"19.100.1049",tanim:"Komple keçiayağı silindir'in 1 saatlik ücreti",birim:"Sa",birimFiyat:1162.88,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1688,no:"19.100.1050",tanim:"2 çift Tanburlu keçiayağı silindir'in 1 saatlik ücreti",birim:"Sa",birimFiyat:1860.61,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1689,no:"19.100.1051",tanim:"3 çift Tanburlu keçiayağı silindir'in 1 saatlik ücreti",birim:"Sa",birimFiyat:2442.05,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1690,no:"19.100.1052",tanim:"Demir merdaneli silindir'in 1 saatlik ücreti",birim:"Sa",birimFiyat:1162.88,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1691,no:"19.100.1053",tanim:"Demir merdaneli silindir'in 1 saatlik ücreti",birim:"Sa",birimFiyat:1324.45,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1692,no:"19.100.1054",tanim:"Lastik tekerlekli silindir'in 1 saatlik ücreti",birim:"Sa",birimFiyat:1162.88,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1693,no:"19.100.1055",tanim:"Küçük eleme tesisi'nin 1 saatlik ücreti",birim:"Sa",birimFiyat:1320.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1694,no:"19.100.1056",tanim:"Fore kazık delgi makinası'nın (300 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:7033.33,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1695,no:"19.100.1057",tanim:"Fore kazık delgi makinası'nın (200 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:3126.89,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1696,no:"19.100.1058",tanim:"Fore kazık delgi makinası'nın (440 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:8929.55,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1697,no:"19.100.1059",tanim:"Otomatik Beton Santrali'nin (1000 Litre kapasiteli, 50 m³/saat) 1 saatlik ücreti",birim:"Sa",birimFiyat:577.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1698,no:"19.100.1060",tanim:"Rotor sistemli BEP 80 M ve benzeri beton pompası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:4205.04,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1699,no:"19.100.1061",tanim:"Kaynak Makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:492.91,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1700,no:"19.100.1062",tanim:"5 kw Jeneratör grubunun 1 saatlik ücreti",birim:"Sa",birimFiyat:757.34,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1701,no:"19.100.1063",tanim:"Klepeli Taş Dubası'nın (125 Ton) 1 saatlik ücreti",birim:"Sa",birimFiyat:610.65,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1702,no:"19.100.1064",tanim:"Klepeli Taş Dubası'nın (400 Ton) 1 saatlik ücreti",birim:"Sa",birimFiyat:899.4,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1703,no:"19.100.1065",tanim:"DevrilirTaş Dubası'ının (300 Ton) 1 saatlik ücreti",birim:"Sa",birimFiyat:899.4,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1704,no:"19.100.1066",tanim:"Klapeli Kum Dubası'ının (300 Ton) 1 saatlik ücreti",birim:"Sa",birimFiyat:899.4,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1705,no:"19.100.1067",tanim:"Klapeli Kum Dubası'ının (2x255 HP, 500 m³) 1 saatlik ücreti",birim:"Sa",birimFiyat:2399.4,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1706,no:"19.100.1068",tanim:"Motorsuz Layter'in (180 m³) 1 saatlik ücreti",birim:"Sa",birimFiyat:1324.4,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1707,no:"19.100.1069",tanim:"Romorkör'ün (116 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:3336.64,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1708,no:"19.100.1070",tanim:"Romorkör'ün (240 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:4544.99,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1709,no:"19.100.1071",tanim:"Romorkör'ün (310 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:4989.42,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1710,no:"19.100.1072",tanim:"Romorkör'ün (525 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:7689.3,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1711,no:"19.100.1073",tanim:"Romorkör'ün (2x300 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:8262.79,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1712,no:"19.100.1074",tanim:"Maçula'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:9418.22,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1713,no:"19.100.1075",tanim:"Çim biçme makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:206.49,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1714,no:"19.100.1076",tanim:"Ziraat işlerinde kullanılan Traktör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:1068.47,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1715,no:"19.100.1077",tanim:"Ziraat işlerinde kullanılan Traktör'ün 1 Saatlik Ücreti",birim:"Sa",birimFiyat:1259.25,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1716,no:"19.100.1078",tanim:"10 litrelik kollu sırt pülverizatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:275.65,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1717,no:"19.100.1079",tanim:"10 litrelik motorlu sırt pülverizatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:314.69,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1718,no:"19.100.1080",tanim:"El ile çekilen 100 litrelik motorlu pülverizatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:561.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1719,no:"19.100.1081",tanim:"Araç ile çekilen 250 litrelik motorlu pülverizatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:1651.68,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1720,no:"19.100.1082",tanim:"Araç ile taşınır 560 litrelik motorlu pülverizatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:1768.39,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1721,no:"19.100.1083",tanim:"Kum Kutusu ve Nozul'un 1 saatlik ücreti",birim:"Sa",birimFiyat:26.77,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1722,no:"19.100.1084",tanim:"Kendi yürür 1200 litrelik motorlu pülverizatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:1150.38,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1723,no:"19.100.1085",tanim:"Karıştırıcı (Blender)'in 1 saatlik ücreti",birim:"Sa",birimFiyat:7.59,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1724,no:"19.100.1086",tanim:"Araç ile taşınır hidrolik düzenli 2200 litrelik motorlu pülverizatör'ün 1 saatlik ücreti",birim:"Sa",birimFiyat:2290.4,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1725,no:"19.100.1087",tanim:"Alüminyum doğrama imalat atelyesinin 1 saatli ücreti",birim:"Sa",birimFiyat:4634.86,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1726,no:"19.100.1088",tanim:"Plastik doğrama imalat atelyesinin 1 saatli ücreti",birim:"Sa",birimFiyat:4603.42,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1727,no:"19.100.1089",tanim:"Demir doğrama imalat atelyesinin 1 saatli ücreti",birim:"Sa",birimFiyat:5235.8,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1728,no:"19.100.1090",tanim:"Tünel kalıp imalat atelyesinin 1 saatlik ücreti",birim:"Sa",birimFiyat:5281.3,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1729,no:"19.100.1093",tanim:"",birim:"Sa",birimFiyat:503.73,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1730,no:"19.100.1094",tanim:"",birim:"Sa",birimFiyat:483.01,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1731,no:"19.100.1095",tanim:"Eksiz Oluk Makinasının 1 saatlik ücreti",birim:"Sa",birimFiyat:465.39,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1732,no:"19.100.1096",tanim:"(Ekskavatör (280 hp) + Tek Depolu Hareketli Kireç Silosu (130 hp) + Karıştırıcı Uç +",birim:"Sa",birimFiyat:15916.26,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1733,no:"19.100.1097",tanim:"(Ekskavatör (280 hp) + Çift Depolu Hareketli Kireç Silosu (130 hp) + Karıştırıcı Uç +",birim:"Sa",birimFiyat:17797.26,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1734,no:"19.100.1098",tanim:"",birim:"Sa",birimFiyat:14614.99,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1735,no:"19.100.1099",tanim:"Zemin Stabilizasyon Makinalarının (Kireç Serme Makinası - 250 hp) bir saatlik ücreti",birim:"Sa",birimFiyat:5569.33,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1736,no:"19.100.1100",tanim:"Sıva makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:749.85,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1737,no:"19.100.1101",tanim:"Mobil Beton Pompasının Bir Saatlik Ücreti (420 HP)",birim:"Sa",birimFiyat:6653.59,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1738,no:"19.100.1102",tanim:"Vinç'in 1 saatlik ücreti",birim:"Sa",birimFiyat:4108.04,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1739,no:"19.100.1103",tanim:"60 Ton kapasiteli mobil vinç'in 1 saatlik ücreti",birim:"Sa",birimFiyat:5309.19,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1740,no:"19.100.1104",tanim:"Mobil vinç'in (60ton, 240 HP) 1 saatlik ücreti",birim:"Sa",birimFiyat:4850.26,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1741,no:"19.100.1105",tanim:"Kule vinçin 1 saatlik ücreti",birim:"Sa",birimFiyat:3590.32,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1742,no:"19.100.1106",tanim:"Paletli delgi makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:3393.51,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1743,no:"19.100.1107",tanim:"Jet grouting ekipmanı ile delgi makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:9662.04,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1744,no:"19.100.1108",tanim:"Mikro tünel sistemi ile boru sürme makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:11814.11,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1745,no:"19.100.1109",tanim:"Mikro tünel sistemi ile boru sürme makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:43083.33,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1746,no:"19.100.1110",tanim:"Matkap'ın 1 saatlik ücreti",birim:"Sa",birimFiyat:369.82,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1747,no:"19.100.1111",tanim:"Demir kesme ve bükme makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:50.47,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1748,no:"19.100.1112",tanim:"Forklift'in 1 saatlik ücreti",birim:"Sa",birimFiyat:885.98,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1749,no:"19.100.1113",tanim:"Mobil vinç'in 1 saatlik ücreti",birim:"Sa",birimFiyat:1402.45,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1750,no:"19.100.1114",tanim:"Çift kompenantlı yalıtım malzemesi dozaj miks makinasının bir saatlik ücreti",birim:"Sa",birimFiyat:930.4,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1751,no:"19.100.1115",tanim:"Grab makinesi'nin 1 saatlik ücreti",birim:"Sa",birimFiyat:15094.73,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1752,no:"19.100.1116",tanim:"Hidrofreze makinesi'nin 1 saatlik ücreti",birim:"Sa",birimFiyat:59196.71,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1753,no:"19.100.1117",tanim:"Bentonit ünitesi ve Desander 1 saatlik ücreti (Grab için)",birim:"Sa",birimFiyat:17460.28,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1754,no:"19.100.1118",tanim:"Bentonit ünitesi ve Desander 1 saatlik ücreti (Hidrofreze için)",birim:"Sa",birimFiyat:23441.37,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1755,no:"19.100.1119",tanim:"Salyangoz pompanın 1 saatlik ücreti",birim:"Sa",birimFiyat:2069.15,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1756,no:"19.100.1120",tanim:"",birim:"Sa",birimFiyat:3900.9,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1757,no:"19.100.1121",tanim:"",birim:"Sa",birimFiyat:417.01,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1758,no:"19.100.1122",tanim:"Bor katkılı selüloz yünü püskürtme makinası'nın 1 saatlik ücreti",birim:"Sa",birimFiyat:472.75,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1759,no:"19.100.2001",tanim:"El arabası ile her cins malzeme ve kayadan başka kazının taşınması (50 metre'ye)",birim:"ton",birimFiyat:133.25,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1760,no:"19.100.2002",tanim:"El arabası ile her cins malzeme ve kayadan başka kazının taşınması (60 metre'ye)",birim:"ton",birimFiyat:159.9,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1761,no:"19.100.2004",tanim:"Makina ile elenmiş, yıkanmış granülometrik kum-çakıl hazırlanması",birim:"m³",birimFiyat:336.24,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1762,no:"19.100.2003",tanim:"Makina ile kum çakıl hazırlanması",birim:"m³",birimFiyat:133.07,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1763,no:"19.100.2005",tanim:"Makina ile elenmiş, yıkanmış granülometrik kum-çakıl hazırlanması",birim:"m³",birimFiyat:279.6,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1764,no:"19.100.2006",tanim:"Makina ile elenmiş, yıkanmış iki tane sınıfı ayrılmış granülometrik kum-çakıl hazırlanması",birim:"m³",birimFiyat:380.74,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1765,no:"19.100.2007",tanim:"Makina ile elenmiş, yıkanmış ince sıva veya derz kumu hazırlanması",birim:"m³",birimFiyat:429.99,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1766,no:"19.100.2008",tanim:"Makina ile elenmiş, yıkanmış ince sıva veya derz kumu hazırlanması",birim:"m³",birimFiyat:281.67,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1767,no:"19.100.2009",tanim:"Demirden basit imalat (kaynaklı - kaynaksız) yapılması",birim:"kg",birimFiyat:237.46,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1768,no:"19.100.2010",tanim:"Demirden basit imalat (kaynaklı - kaynaksız) yapılması",birim:"kg",birimFiyat:233.35,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1769,no:"19.100.2011",tanim:"Demirden basit imalat (kaynaklı - kaynaksız) yapılması",birim:"kg",birimFiyat:236.19,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1770,no:"19.100.2012",tanim:"Galvanizli sacdan Z profilli aşık yapılması",birim:"kg",birimFiyat:104.74,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1771,no:"19.100.2013",tanim:"Demirden basit imalat (kaynaklı - kaynaksız) yapılması",birim:"kg",birimFiyat:251.76,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1772,no:"19.100.2014",tanim:"Vidye kron hazırlanması",birim:"adet",birimFiyat:5758.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1773,no:"19.100.2015",tanim:"Ocakta taş hazırlanması",birim:"m³",birimFiyat:345.39,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1774,no:"19.100.2016",tanim:"Taştan konkasörle kırılmış ve elenmiş 70 mm'ye kadar kırma taş hazırlanması",birim:"m³",birimFiyat:421.95,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1775,no:"19.100.2017",tanim:"Taştan konkasörle kırılmış ve elenmiş 30 mm'ye kadar kırma taş hazırlanması",birim:"m³",birimFiyat:480.43,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1776,no:"19.100.2018",tanim:"Kazıdan taş hazırlanması",birim:"m³",birimFiyat:602.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1777,no:"19.100.2019",tanim:"Toplama suretiyle taş hazırlanması",birim:"m³",birimFiyat:1102.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1778,no:"19.100.2020",tanim:"Ocakta çaplanmış moloz taşı hazırlanması",birim:"m³",birimFiyat:1400.39,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1779,no:"19.100.2021",tanim:"Ocakta, kemer inşaatı için çaplanmış moloz taşı hazırlanması",birim:"m³",birimFiyat:6101.4,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1780,no:"19.100.2022",tanim:"Kazıdan çıkan taşlardan çaplanmış moloz taşı hazırlanması",birim:"m³",birimFiyat:2872.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1781,no:"19.100.2023",tanim:"Kazıdan çıkan taşlardan kemer inşaat için çaplanmış moloz taşı hazırlanması",birim:"m³",birimFiyat:3492.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1782,no:"19.100.2024",tanim:"Ocakta yonu taşı taslağı hazırlanması",birim:"m³",birimFiyat:4240.62,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1783,no:"19.100.2025",tanim:"Ocakta özel yonu taşı taslağı hazırlanması",birim:"m³",birimFiyat:5490.87,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1784,no:"19.100.2026",tanim:"Ocakta hazırlanmış kaba yonu taşı taslağından kaba yonu taşı hazırlanması",birim:"m³",birimFiyat:8612.81,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1785,no:"19.100.2027",tanim:"Kazı taşından hazırlanmış yonu taşı taslağından kaba yonu taşı hazırlanması",birim:"m³",birimFiyat:7640.25,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1786,no:"19.100.2028",tanim:"Ocakta hazırlanmış özel yonu taşı taslağından kaba yonu taşı hazırlanması",birim:"m³",birimFiyat:11717.22,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1787,no:"19.100.2029",tanim:"Kazı taşından hazırlanmış yonu taşı taslağından kaba yonu taşı hazırlanması",birim:"m³",birimFiyat:10655.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1788,no:"19.100.2030",tanim:"Ocakta hazırlanmış yonu taşı taslağından ince yonu taşı hazırlanması",birim:"m³",birimFiyat:13686.87,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1789,no:"19.100.2031",tanim:"Kazı taşından hazırlanmış yonu taşı taslağından ince yonu taşı hazırlanması",birim:"m³",birimFiyat:12639.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1790,no:"19.100.2032",tanim:"Ocakta hazırlanmış özel yonu taşı taslağından özel ince yonu taşı hazırlanması",birim:"m³",birimFiyat:18537.22,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1791,no:"19.100.2033",tanim:"Kazı taşından hazırlanmış özel yonu taşı taslağından özel ince yonu taşı hazırlanması",birim:"m³",birimFiyat:17475.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1792,no:"19.100.2034",tanim:"Ocakta hazırlanmış yonu taşı taslağından kesme taş hazırlanması",birim:"m³",birimFiyat:25384.99,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1793,no:"19.100.2035",tanim:"Kazı taşından hazırlanmış yonu taşı taslağından kesme taş hazırlanması",birim:"m³",birimFiyat:24188.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1794,no:"19.100.2036",tanim:"Yumuşak kesme taş hazırlanması",birim:"m³",birimFiyat:7750.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1795,no:"19.100.2037",tanim:"Doğal taşlarla, alt yüzü kabaca murçlanmış ince yonu kaplama taşı hazırlanması",birim:"m³",birimFiyat:25384.99,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1796,no:"19.100.2038",tanim:"Doğal taşlarla, alt yüzü kabaca tesviye edilmiş ince yonu kaplama taşı hazırlanması",birim:"m³",birimFiyat:28484.99,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1797,no:"19.100.2039",tanim:"Kazı taşından yonu taşı taslağı hazırlanması",birim:"m³",birimFiyat:3492.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1798,no:"19.100.2040",tanim:"Kazı taşından çıkan taşlardan özel yonu taşı taslağı hazırlanması",birim:"m³",birimFiyat:4732.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1799,no:"19.100.2041",tanim:"Ocak taşından (0-0.005 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:71.84,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1800,no:"19.100.2042",tanim:"Kazı taşından (0-0.005 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:22.86,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1801,no:"19.100.2043",tanim:"Ocak taşından (0.005-0.100 ton kategorideki)taşın hazırlanması",birim:"ton",birimFiyat:149.01,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1802,no:"19.100.2044",tanim:"Kazı taşından (0.005-0.100 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:38.06,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1803,no:"19.100.2045",tanim:"Ocak taşından (0-0.250 ton kategorideki)taşın hazırlanması",birim:"ton",birimFiyat:164.82,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1804,no:"19.100.2046",tanim:"Kazı taşından (0-0.250 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:41.88,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1805,no:"19.100.2047",tanim:"Ocak taşından (0.100-0.250 ton kategorideki)taşın hazırlanması",birim:"ton",birimFiyat:170.87,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1806,no:"19.100.2048",tanim:"Kazı taşından (0.100-0.250 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:47.57,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1807,no:"19.100.2049",tanim:"Ocak taşından (0-0.400 ton kategorideki)taşın hazırlanması",birim:"m³",birimFiyat:315.95,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1808,no:"19.100.2050",tanim:"Kazı taşından (0-0.400 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:51.4,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1809,no:"19.100.2051",tanim:"Ocak taşından (0.250-0.400 ton kategorideki)taşın hazırlanması",birim:"ton",birimFiyat:182.07,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1810,no:"19.100.2052",tanim:"Ocak taşından (0.4-2 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:188.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1811,no:"19.100.2053",tanim:"Kazı taşından (0.4-2.0 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:64.75,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1812,no:"19.100.2054",tanim:"Ocak taşından (2-6 ton kategorideki)taşın hazırlanması",birim:"ton",birimFiyat:234.75,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1813,no:"19.100.2055",tanim:"Kazı taşından (2-6 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:85.9,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1814,no:"19.100.2056",tanim:"Ocak taşından (6-15 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:295.29,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1815,no:"19.100.2057",tanim:"Kazı taşından (6-15 ton kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:105.06,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1816,no:"19.100.2058",tanim:"Ocak taşından (15ton'dan yukarı kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:348.87,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1817,no:"19.100.2059",tanim:"Kazı taşından (15ton'dan yukarı kategorideki) taşın hazırlanması",birim:"ton",birimFiyat:134.15,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1818,no:"19.100.2060",tanim:"Ocak artığının ocakta toplanması",birim:"m³",birimFiyat:32.58,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1819,no:"19.100.2061",tanim:"Ocakta kalan kategori fazlası taş",birim:"m³",birimFiyat:142.09,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1820,no:"19.100.2062",tanim:"İş yerinde 1 m³ söndürülmemiş parça kalker kireci hazırlanması",birim:"m³",birimFiyat:2215.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1821,no:"19.100.2063",tanim:"Elle 1 m³ su hazırlanması",birim:"m³",birimFiyat:307.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1822,no:"19.100.2064",tanim:"Motorlu tulumba ile su hazırlanması",birim:"m³",birimFiyat:20.47,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1823,no:"19.100.2065",tanim:"Çelik boru başlarının kaynakla bağlanması",birim:"m",birimFiyat:1152.86,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1824,no:"19.100.2416",tanim:"Çamur harcı hazırlanması",birim:"m³",birimFiyat:1070.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1825,no:"19.100.2426",tanim:"Mozayik harcı yapılması (beyaz çimentolu)",birim:"m³",birimFiyat:4863.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1826,no:"19.100.2432",tanim:"Saten alçı harcı hazırlanması",birim:"m³",birimFiyat:4243.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1827,no:"19.100.2433",tanim:"Perlitli alçı harcı hazırlanması",birim:"m³",birimFiyat:2219.73,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1828,no:"19.100.2434",tanim:"Derz dolgu alçısı harcı hazırlanması",birim:"m³",birimFiyat:4593.95,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1829,no:"19.100.2435",tanim:"Yapıştırma alçısı harcı hazırlanması",birim:"m³",birimFiyat:4593.95,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1830,no:"19.100.2501",tanim:"200 Kg çimento dozlu harç yapılması",birim:"m³",birimFiyat:1779.65,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1831,no:"19.100.2502",tanim:"250 Kg çimento dozlu tesviye harcı yapılması",birim:"m³",birimFiyat:1702.15,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1832,no:"19.100.2503",tanim:"200 Kg çimento dozlu harç yapılması",birim:"m³",birimFiyat:1578.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1833,no:"19.100.2504",tanim:"250 kg çimento dozlu harç yapılması (kargir işlerde)",birim:"m³",birimFiyat:1706.83,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1834,no:"19.100.2505",tanim:"300 Kg çimento dozlu harç yapılması",birim:"m³",birimFiyat:1835.15,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1835,no:"19.100.2506",tanim:"300 Kg çimento dozlu ince harç yapılması",birim:"m³",birimFiyat:1926.15,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1836,no:"19.100.2507",tanim:"350 Kg çimento dozlu harç yapılması",birim:"m³",birimFiyat:1963.48,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1837,no:"19.100.2508",tanim:"350 Kg çimento dozlu ince harç yapılması",birim:"m³",birimFiyat:2054.48,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1838,no:"19.100.2509",tanim:"400 Kg çimento dozlu harç yapılması",birim:"m³",birimFiyat:2091.8,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1839,no:"19.100.2510",tanim:"Kum ve kırmataş ile 400 Kg çimento dozlu harç yapılması",birim:"m³",birimFiyat:2151.3,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1840,no:"19.100.2511",tanim:"400 Kg çimento dozlu ince harç yapılması",birim:"m³",birimFiyat:2182.8,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1841,no:"19.100.2512",tanim:"450 Kg çimento dozlu harç yapılması",birim:"m³",birimFiyat:2220.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1842,no:"19.100.2513",tanim:"450 Kg çimento dozlu ince harç yapılması",birim:"m³",birimFiyat:2311.13,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1843,no:"19.100.2514",tanim:"500 Kg çimento dozlu harç yapılması",birim:"m³",birimFiyat:2348.45,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1844,no:"19.100.2515",tanim:"500 Kg çimento dozlu ince harç yapılması",birim:"m³",birimFiyat:2644.45,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1845,no:"19.100.2517",tanim:"600 dozlu çimento şerbeti",birim:"m³",birimFiyat:1995.0,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1846,no:"19.100.2518",tanim:"500 dozlu çimento şerbeti",birim:"m³",birimFiyat:1777.6,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1847,no:"19.100.2519",tanim:"Kireç harcı yapılması (sönmüş kireç torbalı)",birim:"m³",birimFiyat:1941.51,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1848,no:"19.100.2520",tanim:"Kireç çimento karışımı harç yapılması",birim:"m³",birimFiyat:1941.51,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1849,no:"19.100.2521",tanim:"0.100 m³/250 Kg kireç-çimento karışımı ince harç yapılması(sönmüş kireç torbalı)",birim:"m³",birimFiyat:2032.51,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1850,no:"19.100.2522",tanim:"0.170 m³/200 Kg kireç ve çimento karışımı kaba harç yapılması (sönmüş kireç torbalı)",birim:"m³",birimFiyat:2077.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1851,no:"19.100.2523",tanim:"0.200 m³/150 Kg kireç-çimento karışımı kaba harç yapılması (sönmüş kireç torbalı)",birim:"m³",birimFiyat:1966.88,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1852,no:"19.100.2525",tanim:"Mozayik harcı yapılması",birim:"m³",birimFiyat:3205.5,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1853,no:"19.100.3001",tanim:"",birim:"Sa",birimFiyat:1955.06,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1854,no:"19.100.3002",tanim:"yollar için)",birim:"Sa",birimFiyat:7979.44,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1855,no:"19.100.3003",tanim:"Kayar Kalıplı Beton Finişerinin (Beton Yol İçin) Bir Saatlik Ücreti",birim:"Sa",birimFiyat:37189.8,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
  {id:1856,no:"19.100.3004",tanim:"",birim:"Sa",birimFiyat:7161.67,idare:"Çevre ve Şehircilik",kategori:"Zemin İşleri",kaynak:"2026"},
];

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
  const pozKaynagiRef = useRef({ yil: 2026, ay: null, tip: "Varsayılan" });
  const [pozKaynagi, _setPozKaynagi] = useState({ yil: 2026, ay: null, tip: "Varsayılan" });
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

      {/* 3. Poz ekleme hedef bölümü belirgin göster */}
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
          <div key={poz.id} style={S.pozKart}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
              <span style={{fontFamily:"monospace",fontSize:12,color:"#6366f1",fontWeight:700}}>{poz.no}</span>
              <span style={{...S.chip,backgroundColor:(KAT_RENK[poz.kategori]||"#475569")+"18",color:KAT_RENK[poz.kategori]||"#475569"}}>{poz.kategori}</span>
            </div>
            <div style={{fontSize:13,color:"#cbd5e1",lineHeight:1.5,marginBottom:10}}>{poz.tanim}</div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                <span style={{fontSize:11,color:"#475569",backgroundColor:"#1e293b",padding:"2px 8px",borderRadius:4}}>{poz.idare}</span>
                <span style={{fontSize:12,color:"#475569"}}>/ {poz.birim}</span>
              </div>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <span style={{fontSize:14,fontWeight:700,color:"#10b981"}}>₺{fmt(poz.birimFiyat)}</span>
                {pozEklemeHedef
                  ? <button style={S.birincilBtn} onClick={()=>pozEkle(pozEklemeHedef.projeId,pozEklemeHedef.bolumId,poz.id)}>+ Ekle</button>
                  : <span style={{fontSize:12,color:"#475569"}}>← Bölüm seçin</span>}
              </div>
            </div>
          </div>
        ))}
      </div>
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
