import { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

// ── Sabitler ──────────────────────────────────────────────────────
const IDARELER = ["Tümü","KGM","DSİ","MEB","Çevre ve Şehircilik","Sağlık Bakanlığı","Belediye"];
const KATEGORILER = ["Tümü","Beton İşleri","Demir İşleri","Kaba İnşaat","İnce İnşaat","Tesisat","Elektrik","Zemin İşleri","Nakliye","Diğer"];

const POZLAR_BASLANGIC = [
  { id:1,  no:"16.001/1", tanim:"C20 hazır beton dökülmesi",                  birim:"m³", birimFiyat:2850, idare:"Çevre ve Şehircilik", kategori:"Beton İşleri" },
  { id:2,  no:"16.002/1", tanim:"C25 hazır beton dökülmesi",                  birim:"m³", birimFiyat:3100, idare:"Çevre ve Şehircilik", kategori:"Beton İşleri" },
  { id:3,  no:"21.001",   tanim:"Nervürlü çelik hasır donatı yapılması",      birim:"kg", birimFiyat:42,   idare:"Çevre ve Şehircilik", kategori:"Demir İşleri" },
  { id:4,  no:"KGM-101",  tanim:"Plentmiks temel yapılması",                  birim:"m³", birimFiyat:1250, idare:"KGM",                  kategori:"Zemin İşleri" },
  { id:5,  no:"KGM-202",  tanim:"BSK binder tabakası (yol yüzeyi)",           birim:"ton",birimFiyat:4200, idare:"KGM",                  kategori:"Kaba İnşaat"  },
  { id:6,  no:"DSİ-301",  tanim:"Prefabrik beton kanal kaplama",              birim:"m",  birimFiyat:780,  idare:"DSİ",                  kategori:"Beton İşleri" },
  { id:7,  no:"DSİ-302",  tanim:"Toprak kazısı (1. grup, yumuşak zemin)",     birim:"m³", birimFiyat:145,  idare:"DSİ",                  kategori:"Zemin İşleri" },
  { id:8,  no:"MEB-401",  tanim:"Alçı sıva (iç yüzey)",                       birim:"m²", birimFiyat:95,   idare:"MEB",                  kategori:"İnce İnşaat"  },
  { id:9,  no:"MEB-402",  tanim:"Çift kat yağlı boya (iç cephe)",             birim:"m²", birimFiyat:65,   idare:"MEB",                  kategori:"İnce İnşaat"  },
  { id:10, no:"SAG-501",  tanim:"Alçıpan asma tavan (12.5 mm)",               birim:"m²", birimFiyat:320,  idare:"Sağlık Bakanlığı",     kategori:"İnce İnşaat"  },
  { id:11, no:"BLD-601",  tanim:"Ø100 mm PVC boru döşenmesi (pis su)",        birim:"m",  birimFiyat:185,  idare:"Belediye",             kategori:"Tesisat"      },
  { id:12, no:"BLD-602",  tanim:"Ø150 mm beton boru döşenmesi (yağmur suyu)",birim:"m",  birimFiyat:290,  idare:"Belediye",             kategori:"Tesisat"      },
  { id:13, no:"16.050/3", tanim:"Demirli beton (nervürlü Ø8-Ø32)",            birim:"m³", birimFiyat:3650, idare:"Çevre ve Şehircilik", kategori:"Demir İşleri" },
  { id:14, no:"KGM-110",  tanim:"Kalıp yapımı, söküm ve temizliği",           birim:"m²", birimFiyat:210,  idare:"KGM",                  kategori:"Kaba İnşaat"  },
  { id:15, no:"ELK-701",  tanim:"NYY kablo (4x10 mm²) döşenmesi",            birim:"m",  birimFiyat:155,  idare:"Çevre ve Şehircilik", kategori:"Elektrik"     },
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
    if (!file || file.type !== "application/pdf") {
      setHata("Lütfen PDF dosyası seçin."); setDurum("hata"); return;
    }
    setDosyaAdi(file.name); setDurum("yukleniyor");
    const reader = new FileReader();
    reader.onload = (e) => {
      setPdfBase64(e.target.result.split(",")[1]);
      setDurum("hazir");
    };
    reader.onerror = () => { setHata("Dosya okunamadı."); setDurum("hata"); };
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
