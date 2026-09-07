"""
SQL Veritabanı Şeması ve Bağlantı Yönetimi
"""
import sqlite3
import os
from datetime import datetime
from typing import Optional, List, Dict, Tuple

class Database:
    def __init__(self, db_path: str = "finans_analiz.db"):
        self.db_path = db_path
        self.conn = None
        self.init_database()
    
    def connect(self):
        """Veritabanına bağlan"""
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        return self.conn
    
    def disconnect(self):
        """Veritabanı bağlantısını kapat"""
        if self.conn:
            self.conn.close()
    
    def init_database(self):
        """Veritabanı şemasını oluştur"""
        conn = self.connect()
        cursor = conn.cursor()
        
        # 1. Sektörler Tablosu
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sektorler (
                sektor_id INTEGER PRIMARY KEY AUTOINCREMENT,
                sektor_adi TEXT NOT NULL UNIQUE,
                sektor_grubu TEXT NOT NULL,
                alt_sektor TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 2. Firmalar Tablosu
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS firmalar (
                firma_id INTEGER PRIMARY KEY AUTOINCREMENT,
                firma_kodu TEXT NOT NULL UNIQUE,
                firma_adi TEXT NOT NULL,
                sektor_id INTEGER NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (sektor_id) REFERENCES sektorler(sektor_id)
            )
        ''')
        
        # 3. Finansal Raporlar Tablosu
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
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (firma_id) REFERENCES firmalar(firma_id),
                UNIQUE(firma_id, rapor_tipi, donem_tarihi)
            )
        ''')
        
        # 4. Bilanço Kalemleri Tablosu
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS bilanco_kalemleri (
                kalem_id INTEGER PRIMARY KEY AUTOINCREMENT,
                rapor_id INTEGER NOT NULL,
                kalem_adi TEXT NOT NULL,
                kalem_kodu TEXT,
                kiyas_donem_tutari REAL,
                cari_donem_tutari REAL,
                altinda_mi BOOLEAN DEFAULT 0,
                seq_order INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (rapor_id) REFERENCES finansal_raporlar(rapor_id)
            )
        ''')
        
        # 5. Gelir Tablosu Kalemleri Tablosu
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS gelir_tablosu_kalemleri (
                kalem_id INTEGER PRIMARY KEY AUTOINCREMENT,
                rapor_id INTEGER NOT NULL,
                kalem_adi TEXT NOT NULL,
                kalem_kodu TEXT,
                cari_donem_tutari REAL,
                onceki_donem_tutari REAL,
                cari_donem_3aylik_tutari REAL,
                onceki_donem_3aylik_tutari REAL,
                altinda_mi BOOLEAN DEFAULT 0,
                seq_order INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (rapor_id) REFERENCES finansal_raporlar(rapor_id)
            )
        ''')
        
        # 6. Nakit Akışı Tablosu Kalemleri Tablosu
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS nakitakisi_kalemleri (
                kalem_id INTEGER PRIMARY KEY AUTOINCREMENT,
                rapor_id INTEGER NOT NULL,
                kalem_adi TEXT NOT NULL,
                kalem_kodu TEXT,
                cari_donem_tutari REAL,
                onceki_donem_tutari REAL,
                altinda_mi BOOLEAN DEFAULT 0,
                seq_order INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (rapor_id) REFERENCES finansal_raporlar(rapor_id)
            )
        ''')
        
        # 7. Hesaplanan Metrikler Tablosu
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
        
        # 8. Benchmark Değerleri Tablosu
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
                percentil_90 REAL,
                firma_sayisi INTEGER,
                hesaplama_tarihi TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (sektor_id) REFERENCES sektorler(sektor_id),
                UNIQUE(sektor_id, metrik_adi)
            )
        ''')
        
        # 9. Skor Değerleri Tablosu
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS skor_degerleri (
                skor_id INTEGER PRIMARY KEY AUTOINCREMENT,
                rapor_id INTEGER NOT NULL,
                metrik_adi TEXT NOT NULL,
                skor REAL,
                percentil REAL,
                hesaplama_tarihi TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (rapor_id) REFERENCES finansal_raporlar(rapor_id),
                UNIQUE(rapor_id, metrik_adi)
            )
        ''')
        
        # 10. Veri Senkronizasyon Logu Tablosu
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
        """SELECT sorgusu çalıştır"""
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
        """INSERT/UPDATE/DELETE sorgusu çalıştır"""
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
        except Exception as e:
            conn.rollback()
            conn.close()
            raise e
    
    def get_firma_id(self, firma_kodu: str) -> Optional[int]:
        """Firma kodundan firma_id al"""
        results = self.execute_query(
            "SELECT firma_id FROM firmalar WHERE firma_kodu = ?",
            (firma_kodu,)
        )
        return results[0]['firma_id'] if results else None
    
    def firma_var_mi(self, firma_kodu: str) -> bool:
        """Firma var mı kontrol et"""
        return self.get_firma_id(firma_kodu) is not None
    
    def rapor_var_mi(self, firma_id: int, rapor_tipi: str, donem_tarihi: str) -> bool:
        """Rapor var mı kontrol et"""
        results = self.execute_query(
            "SELECT rapor_id FROM finansal_raporlar WHERE firma_id = ? AND rapor_tipi = ? AND donem_tarihi = ?",
            (firma_id, rapor_tipi, donem_tarihi)
        )
        return len(results) > 0
    
    def get_rapor_id(self, firma_id: int, rapor_tipi: str, donem_tarihi: str) -> Optional[int]:
        """Rapor ID'sini al"""
        results = self.execute_query(
            "SELECT rapor_id FROM finansal_raporlar WHERE firma_id = ? AND rapor_tipi = ? AND donem_tarihi = ?",
            (firma_id, rapor_tipi, donem_tarihi)
        )
        return results[0]['rapor_id'] if results else None
    
    def log_sinkronizasyon(self, dosya_adi: str, firma_kodu: str, rapor_tipi: str, 
                          donem: str, islem_tipi: str, detay: str, basarili: bool):
        """Senkronizasyon logu kaydet"""
        self.execute_insert_update(
            """INSERT INTO sinkronizasyon_logu 
               (excel_dosya_adi, firma_kodu, rapor_tipi, donem, islem_tipi, detay, basarili)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (dosya_adi, firma_kodu, rapor_tipi, donem, islem_tipi, detay, basarili)
        )


class DatabaseHelper:
    """Veritabanı yardımcı işlevleri"""
    
    @staticmethod
    def get_sektor_id(db: Database, sektor_adi: str) -> Optional[int]:
        """Sektör adından sektör ID'sini al"""
        results = db.execute_query(
            "SELECT sektor_id FROM sektorler WHERE sektor_adi = ?",
            (sektor_adi,)
        )
        return results[0]['sektor_id'] if results else None
    
    @staticmethod
    def insert_sektor(db: Database, sektor_adi: str, sektor_grubu: str, alt_sektor: str = None) -> int:
        """Yeni sektör ekle"""
        return db.execute_insert_update(
            "INSERT INTO sektorler (sektor_adi, sektor_grubu, alt_sektor) VALUES (?, ?, ?)",
            (sektor_adi, sektor_grubu, alt_sektor)
        )
    
    @staticmethod
    def insert_firma(db: Database, firma_kodu: str, firma_adi: str, sektor_id: int) -> int:
        """Yeni firma ekle"""
        return db.execute_insert_update(
            "INSERT INTO firmalar (firma_kodu, firma_adi, sektor_id) VALUES (?, ?, ?)",
            (firma_kodu, firma_adi, sektor_id)
        )
    
    @staticmethod
    def get_tum_sektorler(db: Database) -> List[sqlite3.Row]:
        """Tüm sektörleri getir"""
        return db.execute_query("SELECT * FROM sektorler")
