"""
Finansal Analiz Sistemi - Tüm İşlevler Birleştirilmiş
HTML XLS, XLSX Dosyaları Destekler
"""

import sqlite3
import pandas as pd
import numpy as np
import hashlib
import re
from datetime import datetime
from typing import Dict, List, Tuple, Optional
from pathlib import Path
import warnings

warnings.filterwarnings('ignore')

# ============================================================================
# BÖLÜM 1: VERITABANI YÖNETIMI
# ============================================================================

class Database:
    """SQLite Veritabanı Yönetimi"""
    
    def __init__(self, db_path: str = "finans_analiz.db"):
        self.db_path = db_path
        self.conn = None
        self.init_database()
        print(f"✓ Veritabanı hazırlandı: {db_path}")
    
    def connect(self):
        """Veritabanına bağlan"""
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        return self.conn
    
    def disconnect(self):
        """Bağlantıyı kapat"""
        if self.conn:
            self.conn.close()
    
    def init_database(self):
        """Veritabanı şemasını oluştur"""
        conn = self.connect()
        cursor = conn.cursor()
        
        # Sektörler
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sektorler (
                sektor_id INTEGER PRIMARY KEY AUTOINCREMENT,
                sektor_adi TEXT NOT NULL UNIQUE,
                sektor_grubu TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Firmalar
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS firmalar (
                firma_id INTEGER PRIMARY KEY AUTOINCREMENT,
                firma_kodu TEXT NOT NULL UNIQUE,
                firma_adi TEXT NOT NULL,
                sektor_id INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (sektor_id) REFERENCES sektorler(sektor_id)
            )
        ''')
        
        # Finansal Raporlar
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS finansal_raporlar (
                rapor_id INTEGER PRIMARY KEY AUTOINCREMENT,
                firma_id INTEGER NOT NULL,
                rapor_tipi TEXT NOT NULL,
                donem TEXT NOT NULL,
                donem_tarihi DATE NOT NULL,
                sunum_para_birimi TEXT,
                finansal_tablo_niteligi TEXT,
                hash_degeri TEXT UNIQUE,
                source_file TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (firma_id) REFERENCES firmalar(firma_id),
                UNIQUE(firma_id, rapor_tipi, donem_tarihi, finansal_tablo_niteligi)
            )
        ''')
        
        # Bilanço Kalemleri
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS bilanco_kalemleri (
                kalem_id INTEGER PRIMARY KEY AUTOINCREMENT,
                rapor_id INTEGER NOT NULL,
                kalem_adi TEXT NOT NULL,
                cari_donem_tutari REAL,
                kiyas_donem_tutari REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (rapor_id) REFERENCES finansal_raporlar(rapor_id)
            )
        ''')
        
        # Gelir Tablosu Kalemleri
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS gelir_tablosu_kalemleri (
                kalem_id INTEGER PRIMARY KEY AUTOINCREMENT,
                rapor_id INTEGER NOT NULL,
                kalem_adi TEXT NOT NULL,
                cari_donem_tutari REAL,
                onceki_donem_tutari REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (rapor_id) REFERENCES finansal_raporlar(rapor_id)
            )
        ''')
        
        # Hesaplanan Metrikler
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS hesaplanan_metrikler (
                metrik_id INTEGER PRIMARY KEY AUTOINCREMENT,
                rapor_id INTEGER NOT NULL,
                metrik_adi TEXT NOT NULL,
                metrik_degeri REAL,
                hesaplama_tarihi TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (rapor_id) REFERENCES finansal_raporlar(rapor_id),
                UNIQUE(rapor_id, metrik_adi)
            )
        ''')
        
        # Benchmark Değerleri
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS benchmark_degerleri (
                benchmark_id INTEGER PRIMARY KEY AUTOINCREMENT,
                sektor_id INTEGER NOT NULL,
                metrik_adi TEXT NOT NULL,
                ortalama REAL,
                medyan REAL,
                minimum REAL,
                maksimum REAL,
                percentil_25 REAL,
                percentil_75 REAL,
                firma_sayisi INTEGER,
                hesaplama_tarihi TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (sektor_id) REFERENCES sektorler(sektor_id),
                UNIQUE(sektor_id, metrik_adi)
            )
        ''')
        
        # Sinkronizasyon Logu
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sinkronizasyon_logu (
                log_id INTEGER PRIMARY KEY AUTOINCREMENT,
                excel_dosya_adi TEXT NOT NULL,
                firma_kodu TEXT NOT NULL,
                rapor_tipi TEXT NOT NULL,
                donem TEXT NOT NULL,
                islem_tipi TEXT NOT NULL,
                detay TEXT,
                basarili BOOLEAN,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def execute_query(self, query: str, params: Tuple = None) -> List[sqlite3.Row]:
        """SELECT sorgusu"""
        conn = self.connect()
        cursor = conn.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            results = cursor.fetchall()
            conn.close()
            return results
        except Exception as e:
            conn.close()
            raise e
    
    def execute_insert_update(self, query: str, params: Tuple = None) -> int:
        """INSERT/UPDATE/DELETE sorgusu"""
        conn = self.connect()
        cursor = conn.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            conn.commit()
            last_id = cursor.lastrowid
            conn.close()
            return last_id
        except sqlite3.IntegrityError as e:
            conn.close()
            return None
        except Exception as e:
            conn.close()
            raise e
    
    def get_firma_id(self, firma_kodu: str) -> Optional[int]:
        """Firma ID'sini al"""
        results = self.execute_query(
            "SELECT firma_id FROM firmalar WHERE firma_kodu = ?",
            (firma_kodu,)
        )
        return results[0]['firma_id'] if results else None
    
    def get_rapor_id(self, firma_id: int, rapor_tipi: str, donem_tarihi: str, finansal_tablo_niteligi: str) -> Optional[int]:
        """Rapor ID'sini al"""
        results = self.execute_query(
            """SELECT rapor_id FROM finansal_raporlar 
               WHERE firma_id = ? AND rapor_tipi = ? AND donem_tarihi = ? AND finansal_tablo_niteligi = ?""",
            (firma_id, rapor_tipi, donem_tarihi, finansal_tablo_niteligi)
        )
        return results[0]['rapor_id'] if results else None


# ============================================================================
# BÖLÜM 2: EXCEL OKUMA (HTML XLS VE XLSX DESTEĞİ)
# ============================================================================

class ExcelReader:
    """Excel dosyalarından veri okuma - HTML XLS, XLSX destekler"""
    
    def __init__(self, dosya_yolu: str):
        self.dosya_yolu = dosya_yolu
        self.dosya_adi = Path(dosya_yolu).name
        self.uzanti = Path(dosya_yolu).suffix.lower()
        print(f"📄 Dosya açılıyor: {self.dosya_adi} ({self.uzanti})")
    
    def sektorler_xlsx_oku(self) -> Dict[str, List[Dict]]:
        """sektorler.xlsx dosyasından sektör-firma bilgilerini oku"""
        try:
            # Tüm sayfaları oku
            xls = pd.ExcelFile(self.dosya_yolu)
            
            sektorler_firmalar = {}
            
            for sheet_name in xls.sheet_names:
                if sheet_name.startswith('Sheet'):
                    continue
                
                df = pd.read_excel(self.dosya_yolu, sheet_name=sheet_name)
                
                # Sütunları temizle
                df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
                
                firmalar = []
                
                for idx, row in df.iterrows():
                    # Kod ve firma adını çıkar
                    if len(row) >= 3:
                        kod = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else None
                        adi = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else None
                        
                        if kod and kod != 'Kod' and kod != ';;' and adi and adi != 'Kayıt Bulunamadı':
                            firmalar.append({
                                'kod': kod,
                                'adi': adi
                            })
                
                if firmalar:
                    sektorler_firmalar[sheet_name] = firmalar
            
            print(f"✓ {len(sektorler_firmalar)} sektörden {sum(len(f) for f in sektorler_firmalar.values())} firma okundu")
            return sektorler_firmalar
        
        except Exception as e:
            print(f"✗ Sektör okuma hatası: {e}")
            return {}
    
    def finansal_tablolar_oku(self) -> Dict[str, List[Dict]]:
        """HTML XLS dosyasından finansal tabloları oku"""
        try:
            # Tüm tabloları oku
            tables = pd.read_html(self.dosya_yolu)
            
            finansal_veriler = {}
            
            for idx, table in enumerate(tables):
                if table.empty or len(table) < 2:
                    continue
                
                # Firma adı ve periyot bilgisini çıkar
                firma_adi, periyot = self._meta_bilgiler_cikar(table)
                
                if firma_adi:
                    key = f"{firma_adi}_{periyot}"
                    finansal_veriler[key] = {
                        'firma_adi': firma_adi,
                        'periyot': periyot,
                        'tablo_no': idx,
                        'data': table
                    }
            
            print(f"✓ {len(finansal_veriler)} finansal tablo okundu")
            return finansal_veriler
        
        except Exception as e:
            print(f"✗ Finansal tablo okuma hatası: {e}")
            return {}
    
    def _meta_bilgiler_cikar(self, df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
        """Tablo içinden firma adı ve periyot bilgisini çıkar"""
        firma_adi = None
        periyot = None
        
        # İlk 10 satırı kontrol et
        for idx, row in df.head(10).iterrows():
            if pd.notna(row.iloc[0]):
                satir_str = str(row.iloc[0]).lower()
                
                if 'ünvan' in satir_str or 'a.ş' in satir_str or 'ltd' in satir_str:
                    if pd.notna(row.iloc[1]):
                        firma_adi = str(row.iloc[1]).strip()
                
                if 'periyot' in satir_str and pd.notna(row.iloc[1]):
                    periyot = str(row.iloc[1]).strip()
        
        return firma_adi, periyot or "2"


# ============================================================================
# BÖLÜM 3: AKILLI VERI SENKRONIZASYONU
# ============================================================================

class DataSync:
    """Akıllı veri senkronizasyonu - mükerrer kontrol ve güncelleme"""
    
    def __init__(self, db: Database):
        self.db = db
    
    @staticmethod
    def hash_hesapla(veri: str) -> str:
        """Veri için hash hesapla"""
        return hashlib.md5(str(veri).encode()).hexdigest()
    
    def sektor_firmalar_yukle(self, sektorler_firmalar: Dict[str, List[Dict]]):
        """Sektör ve firmaları veritabanına yükle"""
        yuklu_sayac = 0
        
        for sektor_adi, firmalar in sektorler_firmalar.items():
            # Sektörü ekle veya getir
            sektor_result = self.db.execute_query(
                "SELECT sektor_id FROM sektorler WHERE sektor_adi = ?",
                (sektor_adi,)
            )
            
            if sektor_result:
                sektor_id = sektor_result[0]['sektor_id']
            else:
                sektor_id = self.db.execute_insert_update(
                    "INSERT INTO sektorler (sektor_adi, sektor_grubu) VALUES (?, ?)",
                    (sektor_adi, sektor_adi)
                )
            
            # Firmaları ekle
            for firma in firmalar:
                firma_id = self.db.execute_insert_update(
                    "INSERT OR IGNORE INTO firmalar (firma_kodu, firma_adi, sektor_id) VALUES (?, ?, ?)",
                    (firma['kod'], firma['adi'], sektor_id)
                )
                if firma_id:
                    yuklu_sayac += 1
        
        print(f"✓ {yuklu_sayac} firma yüklendi")
        return yuklu_sayac
    
    def finansal_rapor_yukle(self, firma_kodu: str, rapor_tipi: str, donem: str, 
                            donem_tarihi: str, finansal_tablo_niteligi: str,
                            bilanco_kalemleri: List[Dict] = None,
                            gelir_kalemleri: List[Dict] = None,
                            source_file: str = None) -> bool:
        """Finansal raporu veritabanına yükle"""
        
        try:
            # Firma ID'sini al
            firma_id = self.db.get_firma_id(firma_kodu)
            if not firma_id:
                print(f"✗ Firma bulunamadı: {firma_kodu}")
                return False
            
            # Hash hesapla
            veri_str = f"{firma_kodu}_{donem_tarihi}_{bilanco_kalemleri}_{gelir_kalemleri}"
            hash_val = self.hash_hesapla(veri_str)
            
            # Varolan raporu kontrol et
            existing_rapor = self.db.execute_query(
                """SELECT rapor_id, hash_degeri FROM finansal_raporlar 
                   WHERE firma_id = ? AND rapor_tipi = ? AND donem_tarihi = ? AND finansal_tablo_niteligi = ?""",
                (firma_id, rapor_tipi, donem_tarihi, finansal_tablo_niteligi)
            )
            
            if existing_rapor:
                existing_hash = existing_rapor[0]['hash_degeri']
                if existing_hash == hash_val:
                    print(f"⊘ Mükerrer rapor atlandı: {firma_kodu} - {donem_tarihi}")
                    return False
                else:
                    # Güncelle
                    rapor_id = existing_rapor[0]['rapor_id']
                    self.db.execute_insert_update(
                        """UPDATE finansal_raporlar SET hash_degeri = ?, updated_at = CURRENT_TIMESTAMP 
                           WHERE rapor_id = ?""",
                        (hash_val, rapor_id)
                    )
                    # Eski kalemleri sil
                    self.db.execute_insert_update(
                        "DELETE FROM bilanco_kalemleri WHERE rapor_id = ?",
                        (rapor_id,)
                    )
                    self.db.execute_insert_update(
                        "DELETE FROM gelir_tablosu_kalemleri WHERE rapor_id = ?",
                        (rapor_id,)
                    )
                    print(f"↻ Rapor güncellendi: {firma_kodu} - {donem_tarihi}")
            else:
                # Yeni rapor ekle
                rapor_id = self.db.execute_insert_update(
                    """INSERT INTO finansal_raporlar 
                       (firma_id, rapor_tipi, donem, donem_tarihi, finansal_tablo_niteligi, hash_degeri, source_file)
                       VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (firma_id, rapor_tipi, donem, donem_tarihi, finansal_tablo_niteligi, hash_val, source_file)
                )
                print(f"✓ Yeni rapor eklendi: {firma_kodu} - {donem_tarihi}")
            
            # Bilanço kalemlerini ekle
            if bilanco_kalemleri:
                for kalem in bilanco_kalemleri:
                    self.db.execute_insert_update(
                        """INSERT INTO bilanco_kalemleri 
                           (rapor_id, kalem_adi, cari_donem_tutari, kiyas_donem_tutari)
                           VALUES (?, ?, ?, ?)""",
                        (rapor_id, kalem['kalem_adi'], kalem.get('cari_donem_tutari'), kalem.get('kiyas_donem_tutari'))
                    )
            
            # Gelir tablosu kalemlerini ekle
            if gelir_kalemleri:
                for kalem in gelir_kalemleri:
                    self.db.execute_insert_update(
                        """INSERT INTO gelir_tablosu_kalemleri 
                           (rapor_id, kalem_adi, cari_donem_tutari, onceki_donem_tutari)
                           VALUES (?, ?, ?, ?)""",
                        (rapor_id, kalem['kalem_adi'], kalem.get('cari_donem_tutari'), kalem.get('onceki_donem_tutari'))
                    )
            
            return True
        
        except Exception as e:
            print(f"✗ Rapor yükleme hatası: {e}")
            return False


# ============================================================================
# BÖLÜM 4: FINANSAL METRİKLER HESAPLAMA
# ============================================================================

class MetrikHesapla:
    """Finansal metrikleri hesapla"""
    
    def __init__(self, db: Database):
        self.db = db
    
    def tum_metrikleri_hesapla(self, rapor_id: int):
        """Bir rapor için tüm metrikleri hesapla"""
        
        # Bilanço verilerini getir
        bilanco = self._bilanco_ver_dict(rapor_id)
        
        # Gelir tablosu verilerini getir
        gelir = self._gelir_tablosu_ver_dict(rapor_id)
        
        metrikler = {}
        
        # Likidite Oranları
        if 'dönen varlıklar' in bilanco and 'kısa vadeli yükümlülükler' in bilanco:
            try:
                metrikler['cari_oran'] = bilanco['dönen varlıklar'] / bilanco['kısa vadeli yükümlülükler']
            except:
                pass
        
        # Karlılık Oranları
        if 'net dönem kârı' in gelir and 'toplam varlıklar' in bilanco:
            try:
                metrikler['ROA'] = (gelir['net dönem kârı'] / bilanco['toplam varlıklar']) * 100
            except:
                pass
        
        if 'net dönem kârı' in gelir and 'toplam özkaynak' in bilanco:
            try:
                metrikler['ROE'] = (gelir['net dönem kârı'] / bilanco['toplam özkaynak']) * 100
            except:
                pass
        
        # Kâr Marjları
        if 'hasılat' in gelir and gelir['hasılat'] != 0:
            if 'brüt kâr' in gelir:
                try:
                    metrikler['brut_kar_marji'] = (gelir['brüt kâr'] / gelir['hasılat']) * 100
                except:
                    pass
            
            if 'net dönem kârı' in gelir:
                try:
                    metrikler['net_kar_marji'] = (gelir['net dönem kârı'] / gelir['hasılat']) * 100
                except:
                    pass
        
        # Borç Oranları
        if 'toplam yükümlülükler' in bilanco and 'toplam varlıklar' in bilanco:
            try:
                metrikler['borç_orani'] = (bilanco['toplam yükümlülükler'] / bilanco['toplam varlıklar']) * 100
            except:
                pass
        
        # Veritabanına kaydet
        for metrik_adi, metrik_degeri in metrikler.items():
            try:
                self.db.execute_insert_update(
                    """INSERT OR REPLACE INTO hesaplanan_metrikler 
                       (rapor_id, metrik_adi, metrik_degeri)
                       VALUES (?, ?, ?)""",
                    (rapor_id, metrik_adi, metrik_degeri)
                )
            except:
                pass
        
        return metrikler
    
    def _bilanco_ver_dict(self, rapor_id: int) -> Dict[str, float]:
        """Bilanço verilerini dictionary olarak getir"""
        results = self.db.execute_query(
            "SELECT kalem_adi, cari_donem_tutari FROM bilanco_kalemleri WHERE rapor_id = ?",
            (rapor_id,)
        )
        
        bilanco = {}
        for row in results:
            kalem = str(row['kalem_adi']).lower()
            tutari = row['cari_donem_tutari']
            
            if tutari is not None:
                if 'dönen' in kalem:
                    bilanco['dönen varlıklar'] = tutari
                elif 'duran' in kalem:
                    bilanco['duran varlıklar'] = tutari
                elif 'toplam varlık' in kalem:
                    bilanco['toplam varlıklar'] = tutari
                elif 'kısa vadeli' in kalem and 'yükümlülük' in kalem:
                    bilanco['kısa vadeli yükümlülükler'] = tutari
                elif 'uzun vadeli' in kalem and 'yükümlülük' in kalem:
                    bilanco['uzun vadeli yükümlülükler'] = tutari
                elif 'toplam yükümlülük' in kalem:
                    bilanco['toplam yükümlülükler'] = tutari
                elif 'toplam özkaynak' in kalem:
                    bilanco['toplam özkaynak'] = tutari
        
        return bilanco
    
    def _gelir_tablosu_ver_dict(self, rapor_id: int) -> Dict[str, float]:
        """Gelir tablosu verilerini dictionary olarak getir"""
        results = self.db.execute_query(
            "SELECT kalem_adi, cari_donem_tutari FROM gelir_tablosu_kalemleri WHERE rapor_id = ?",
            (rapor_id,)
        )
        
        gelir = {}
        for row in results:
            kalem = str(row['kalem_adi']).lower()
            tutari = row['cari_donem_tutari']
            
            if tutari is not None:
                if 'hasılat' in kalem:
                    gelir['hasılat'] = tutari
                elif 'brüt kâr' in kalem:
                    gelir['brüt kâr'] = tutari
                elif 'net dönem kârı' in kalem:
                    gelir['net dönem kârı'] = tutari
        
        return gelir


# ============================================================================
# BÖLÜM 5: BENCHMARK VE SKOR HESAPLAMA
# ============================================================================

class BenchmarkHesapla:
    """Sektör benchmark ve skor hesaplama"""
    
    def __init__(self, db: Database):
        self.db = db
    
    def sektor_benchmark_hesapla(self, sektor_id: int):
        """Sektör için benchmark değerlerini hesapla"""
        
        # Sektördeki tüm metrikleri getir
        results = self.db.execute_query(
            """SELECT DISTINCT m.metrik_adi, m.metrik_degeri
               FROM hesaplanan_metrikler m
               JOIN finansal_raporlar fr ON m.rapor_id = fr.rapor_id
               JOIN firmalar f ON fr.firma_id = f.firma_id
               WHERE f.sektor_id = ? AND m.metrik_degeri IS NOT NULL""",
            (sektor_id,)
        )
        
        metrik_gruplar = {}
        for row in results:
            metrik = row['metrik_adi']
            deger = row['metrik_degeri']
            
            if metrik not in metrik_gruplar:
                metrik_gruplar[metrik] = []
            metrik_gruplar[metrik].append(deger)
        
        # Her metrik için istatistik hesapla
        for metrik_adi, degerler in metrik_gruplar.items():
            if not degerler:
                continue
            
            degerler = [d for d in degerler if d is not None and not np.isnan(d)]
            
            if degerler:
                benchmark = {
                    'ortalama': float(np.mean(degerler)),
                    'medyan': float(np.median(degerler)),
                    'minimum': float(np.min(degerler)),
                    'maksimum': float(np.max(degerler)),
                    'percentil_25': float(np.percentile(degerler, 25)),
                    'percentil_75': float(np.percentile(degerler, 75)),
                    'firma_sayisi': len(degerler)
                }
                
                self.db.execute_insert_update(
                    """INSERT OR REPLACE INTO benchmark_degerleri
                       (sektor_id, metrik_adi, ortalama, medyan, minimum, maksimum, percentil_25, percentil_75, firma_sayisi)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (sektor_id, metrik_adi, benchmark['ortalama'], benchmark['medyan'],
                     benchmark['minimum'], benchmark['maksimum'], benchmark['percentil_25'],
                     benchmark['percentil_75'], benchmark['firma_sayisi'])
                )
        
        print(f"✓ Sektör {sektor_id} için {len(metrik_gruplar)} metrik hesaplandı")
    
    def firma_skor_hesapla(self, rapor_id: int):
        """Firma raporuna skor ver"""
        
        # Rapor bilgilerini getir
        rapor = self.db.execute_query(
            "SELECT firma_id FROM finansal_raporlar WHERE rapor_id = ?",
            (rapor_id,)
        )[0]
        
        firma_id = rapor['firma_id']
        
        # Firmayı sektor bilgisini getir
        firma = self.db.execute_query(
            "SELECT sektor_id FROM firmalar WHERE firma_id = ?",
            (firma_id,)
        )[0]
        
        sektor_id = firma['sektor_id']
        
        # Firmakın metriklerini getir
        metrikler = self.db.execute_query(
            "SELECT metrik_adi, metrik_degeri FROM hesaplanan_metrikler WHERE rapor_id = ?",
            (rapor_id,)
        )
        
        # Her metrik için skor hesapla
        for metrik in metrikler:
            metrik_adi = metrik['metrik_adi']
            metrik_degeri = metrik['metrik_degeri']
            
            if metrik_degeri is None:
                continue
            
            # Benchmark getir
            benchmark = self.db.execute_query(
                """SELECT percentil_25, percentil_75, ortalama FROM benchmark_degerleri
                   WHERE sektor_id = ? AND metrik_adi = ?""",
                (sektor_id, metrik_adi)
            )
            
            if not benchmark:
                continue
            
            bench = benchmark[0]
            
            # Skor hesapla (0-100)
            if metrik_degeri <= bench['percentil_25']:
                skor = 25
            elif metrik_degeri <= bench['ortalama']:
                skor = 50
            elif metrik_degeri <= bench['percentil_75']:
                skor = 75
            else:
                skor = 100
            
            # Yüzdelik hesapla
            if metrik_degeri <= bench['percentil_25']:
                percentil = 25
            elif metrik_degeri <= bench['percentil_75']:
                percentil = 50
            else:
                percentil = 75
            
            self.db.execute_insert_update(
                """INSERT OR REPLACE INTO skor_degerleri
                   (rapor_id, metrik_adi, skor, percentil)
                   VALUES (?, ?, ?, ?)""",
                (rapor_id, metrik_adi, skor, percentil)
            )


# ============================================================================
# BÖLÜM 6: RAPOR ÜRETME
# ============================================================================

class RaporUret:
    """Excel raporları oluştur"""
    
    def __init__(self, db: Database):
        self.db = db
    
    def ozet_rapor_uret(self, output_path: str = "finans_ozet_rapor.xlsx"):
        """Özet rapor oluştur"""
        
        try:
            # Tüm firmalar ve metrikleri getir
            results = self.db.execute_query(
                """SELECT f.firma_adi, f.firma_kodu, s.sektor_adi, 
                          m.metrik_adi, m.metrik_degeri, sk.skor, sk.percentil
                   FROM firmalar f
                   JOIN sektorler s ON f.sektor_id = s.sektor_id
                   LEFT JOIN hesaplanan_metrikler m ON f.firma_id = (
                       SELECT firma_id FROM finansal_raporlar WHERE rapor_id = m.rapor_id LIMIT 1
                   )
                   LEFT JOIN skor_degerleri sk ON m.metrik_id = sk.metrik_id"""
            )
            
            # DataFrame oluştur
            data = []
            for row in results:
                data.append({
                    'Firma': row['firma_adi'],
                    'Kod': row['firma_kodu'],
                    'Sektör': row['sektor_adi'],
                    'Metrik': row['metrik_adi'],
                    'Değer': row['metrik_degeri'],
                    'Skor': row['skor'],
                    'Yüzdelik': row['percentil']
                })
            
            if not data:
                print("✗ Rapor oluşturmak için veri yok")
                return
            
            df = pd.DataFrame(data)
            df.to_excel(output_path, index=False, sheet_name='Özet')
            
            print(f"✓ Özet rapor oluşturuldu: {output_path}")
        
        except Exception as e:
            print(f"✗ Rapor üretme hatası: {e}")


# ============================================================================
# BÖLÜM 7: KULLANIM ÖRNEĞİ
# ============================================================================

class FinansSistemi:
    """Ana sistem sınıfı"""
    
    def __init__(self, db_path: str = "finans_analiz.db"):
        self.db = Database(db_path)
        self.sync = DataSync(self.db)
        self.metrik = MetrikHesapla(self.db)
        self.benchmark = BenchmarkHesapla(self.db)
        self.rapor = RaporUret(self.db)
    
    def sektorler_yukle(self, xlsx_dosya: str):
        """Sektör ve firma listesi yükle"""
        print("\n" + "="*60)
        print("SEKTÖR VE FIRMA YÜKLEMESİ")
        print("="*60)
        
        reader = ExcelReader(xlsx_dosya)
        sektorler_firmalar = reader.sektorler_xlsx_oku()
        
        if sektorler_firmalar:
            self.sync.sektir_firmalar_yukle(sektorler_firmalar)
    
    def finansal_tablolar_yukle(self, xls_dosya: str):
        """Finansal tabloları yükle"""
        print("\n" + "="*60)
        print("FİNANSAL TABLOLAR YÜKLEMESİ")
        print("="*60)
        
        reader = ExcelReader(xls_dosya)
        finansal_veriler = reader.finansal_tablolar_oku()
        
        # Her tabloyu işle
        for key, veri in finansal_veriler.items():
            try:
                firma_adi = veri['firma_adi']
                
                # Firma kodunu addan çıkar (simplistic approach)
                firma_kodu = firma_adi.split()[0] if firma_adi else None
                
                if not firma_kodu:
                    continue
                
                # Bilanço, gelir tablosu vb. kalemleri çıkar
                bilanco_kalemleri = self._bilanco_cikar(veri['data'])
                gelir_kalemleri = self._gelir_cikar(veri['data'])
                
                # Veritabanına yükle
                self.sync.finansal_rapor_yukle(
                    firma_kodu=firma_kodu,
                    rapor_tipi='FINANSAL RAPOR',
                    donem=veri['periyot'],
                    donem_tarihi='2026-06-30',
                    finansal_tablo_niteligi='KONSOLIDE',
                    bilanco_kalemleri=bilanco_kalemleri,
                    gelir_kalemleri=gelir_kalemleri,
                    source_file=xls_dosya
                )
            except Exception as e:
                print(f"✗ {key} işleme hatası: {e}")
    
    def metrikleri_hesapla_ve_skor_ver(self):
        """Tüm raporlar için metrikleri hesapla ve skor ver"""
        print("\n" + "="*60)
        print("METRİK HESAPLAMA VE SKOR VERİLMESİ")
        print("="*60)
        
        # Tüm raporları getir
        raporlar = self.db.execute_query(
            "SELECT rapor_id, firma_id FROM finansal_raporlar"
        )
        
        for rapor in raporlar:
            rapor_id = rapor['rapor_id']
            
            # Metrikleri hesapla
            self.metrik.tum_metrikleri_hesapla(rapor_id)
        
        # Sektörler için benchmark hesapla
        sektorler = self.db.execute_query(
            "SELECT sektor_id FROM sektorler"
        )
        
        for sektor in sektorler:
            sektor_id = sektor['sektor_id']
            self.benchmark.sektor_benchmark_hesapla(sektor_id)
        
        # Firmalar için skor ver
        for rapor in raporlar:
            self.benchmark.firma_skor_hesapla(rapor['rapor_id'])
        
        print("✓ Metrikler hesaplandı ve skorlar verildi")
    
    def ozet_rapor_uret(self):
        """Özet rapor oluştur"""
        print("\n" + "="*60)
        print("ÖZET RAPOR ÜRETME")
        print("="*60)
        
        self.rapor.ozet_rapor_uret()
    
    def _bilanco_cikar(self, df: pd.DataFrame) -> List[Dict]:
        """DataFrame'den bilanço kalemlerini çıkar"""
        kalemleri = []
        
        for idx, row in df.iterrows():
            if len(row) >= 2:
                kalem_adi = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
                
                if kalem_adi and len(kalem_adi) > 2:
                    cari_tutari = None
                    kiyas_tutari = None
                    
                    try:
                        if pd.notna(row.iloc[1]):
                            cari_tutari = float(row.iloc[1])
                    except:
                        pass
                    
                    try:
                        if len(row) > 2 and pd.notna(row.iloc[2]):
                            kiyas_tutari = float(row.iloc[2])
                    except:
                        pass
                    
                    if cari_tutari or kiyas_tutari:
                        kalemleri.append({
                            'kalem_adi': kalem_adi,
                            'cari_donem_tutari': cari_tutari,
                            'kiyas_donem_tutari': kiyas_tutari
                        })
        
        return kalemleri
    
    def _gelir_cikar(self, df: pd.DataFrame) -> List[Dict]:
        """DataFrame'den gelir tablosu kalemlerini çıkar"""
        kalemleri = []
        
        for idx, row in df.iterrows():
            if len(row) >= 2:
                kalem_adi = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else None
                
                if kalem_adi and len(kalem_adi) > 2:
                    cari_tutari = None
                    onceki_tutari = None
                    
                    try:
                        if pd.notna(row.iloc[1]):
                            cari_tutari = float(row.iloc[1])
                    except:
                        pass
                    
                    try:
                        if len(row) > 2 and pd.notna(row.iloc[2]):
                            onceki_tutari = float(row.iloc[2])
                    except:
                        pass
                    
                    if cari_tutari or onceki_tutari:
                        kalemleri.append({
                            'kalem_adi': kalem_adi,
                            'cari_donem_tutari': cari_tutari,
                            'onceki_donem_tutari': onceki_tutari
                        })
        
        return kalemleri


# ============================================================================
# MAIN - KULLANIM ÖRNEĞİ
# ============================================================================

if __name__ == "__main__":
    print("\n" + "="*70)
    print("     FINANSAL ANALİZ SİSTEMİ - BAŞLANGIÇ")
    print("="*70 + "\n")
    
    # Sistem oluştur
    sistem = FinansSistemi()
    
    # ADIM 1: Sektörler ve firmalar yükle
    try:
        sistem.sektorler_yukle("sektorler.xlsx")
    except FileNotFoundError:
        print("✗ sektorler.xlsx bulunamadı. Lütfen dosyayı kontrol edin.")
    except Exception as e:
        print(f"✗ Sektör yükleme hatası: {e}")
    
    # ADIM 2: Finansal tablolar yükle
    try:
        sistem.finansal_tablolar_yukle("kap.xls")
    except FileNotFoundError:
        print("✗ kap.xls bulunamadı. Lütfen dosyayı kontrol edin.")
    except Exception as e:
        print(f"✗ Finansal tablo yükleme hatası: {e}")
    
    # ADIM 3: Metrikleri hesapla ve skor ver
    try:
        sistem.metrikleri_hesapla_ve_skor_ver()
    except Exception as e:
        print(f"✗ Metrik hesaplama hatası: {e}")
    
    # ADIM 4: Özet rapor oluştur
    try:
        sistem.ozet_rapor_uret()
    except Exception as e:
        print(f"✗ Rapor üretme hatası: {e}")
    
    print("\n" + "="*70)
    print("     IŞLEM TAMAMLANDI")
    print("="*70 + "\n")
    
    print("📊 Veritabanı: finans_analiz.db")
    print("📈 Rapor: finans_ozet_rapor.xlsx")
    print("\n✓ Sistem hazırlandı ve veriler yüklendi.\n")
