# MetrajX — Metraj Hesaplama Uygulaması

Profesyonel metraj cetveli, yaklaşık maliyet ve poz kütüphanesi uygulaması.

## Kurulum

### Gereksinimler
- [Node.js](https://nodejs.org) (v18 veya üstü)

### Bilgisayarda Çalıştırma

```bash
# 1. Bağımlılıkları yükle
npm install

# 2. Geliştirme sunucusunu başlat
npm run dev
```

Tarayıcıda `http://localhost:5173` adresini aç.

### Vercel ile Deploy (Ücretsiz)

1. [vercel.com](https://vercel.com) adresine git
2. GitHub hesabınla giriş yap
3. Bu klasörü GitHub'a yükle (yeni repo oluştur)
4. Vercel'de "New Project" → GitHub reposunu seç → Deploy
5. Otomatik URL alırsın (ör: `metrajx.vercel.app`)

### Netlify ile Deploy (Ücretsiz)

1. `npm run build` komutu çalıştır → `dist` klasörü oluşur
2. [netlify.com](https://netlify.com) → "Deploy manually"
3. `dist` klasörünü sürükle-bırak
4. URL alırsın

## Özellikler

- 🏗 Proje / Bölüm / Poz hiyerarşisi
- 📐 Metraj cetveli (boyut girişi + minha/düşme)
- 📚 Poz kütüphanesi (PDF import ile güncelleme)
- 📊 Özet & KDV hesabı
- 📥 Excel çıktısı (7 rapor tipi)
- 💾 Otomatik veri kaydetme
