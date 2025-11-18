"""
Outlook Email Address Scanner

A comprehensive tool to extract email addresses from Outlook mailboxes.
Scans sent items and archive folders to collect unique email addresses.

Author: Baran Can Balta
Email: barancanbalta@outlook.com
"""

from __future__ import annotations

import logging
import os
import time
import contextlib
from datetime import datetime, timedelta
from email.utils import getaddresses
from pathlib import Path
from typing import Dict, Optional, Sequence, Set
from dataclasses import dataclass, field

import pandas as pd
import win32com.client

# ============================================================================
# SABITLER
# ============================================================================
PROP_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_ARCHIVE = 109
PROP_TO = "http://schemas.microsoft.com/mapi/proptag/0x0E04001E"
PROP_CC = "http://schemas.microsoft.com/mapi/proptag/0x0E03001E"
PROP_BCC = "http://schemas.microsoft.com/mapi/proptag/0x0E1D001E"
ADDRESS_COLUMNS: Sequence[str] = (PROP_TO, PROP_CC, PROP_BCC)
ARCHIVE_ALIASES = {"archive", "arşiv", "archive items"}

# Checkpoint parametreleri
CHECKPOINT_INTERVAL = 1000  # Her 1000 kayıtta bir backup
PROGRESS_LOG_INTERVAL = 200  # Her 200 kayıtta bir log


# ============================================================================
# VERİ YAPILARI
# ============================================================================
@dataclass
class ScanStats:
    """Tarama istatistikleri"""
    total_emails_scanned: int = 0
    total_addresses_found: int = 0
    errors_encountered: int = 0
    folders_scanned: int = 0
    start_time: float = field(default_factory=time.perf_counter)
    
    def elapsed(self) -> float:
        return time.perf_counter() - self.start_time
    
    def summary(self) -> str:
        return (
            f"Toplam Mail: {self.total_emails_scanned} | "
            f"Toplam Adres: {self.total_addresses_found} | "
            f"Hata: {self.errors_encountered} | "
            f"Klasör: {self.folders_scanned} | "
            f"Süre: {self.elapsed():.1f}sn"
        )


# ============================================================================
# LOGGING YAPISI
# ============================================================================
def configure_logging(verbose: bool = False, log_file: Optional[Path] = None) -> None:
    """Gelişmiş logging yapılandırması"""
    level = logging.DEBUG if verbose else logging.INFO
    
    handlers = [logging.StreamHandler()]
    
    if log_file:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        handlers.append(file_handler)
    
    logging.basicConfig(
        level=level,
        format="%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%H:%M:%S",
        handlers=handlers,
        force=True  # Önceki config'i override et
    )


# ============================================================================
# E-POSTA NORMALIZASYON VE ÇÖZÜMLEME
# ============================================================================
def normalize_email(addr: Optional[str]) -> Optional[str]:
    """E-posta adresini normalize et ve geçerli olup olmadığını kontrol et"""
    if not addr:
        return None
    
    addr = addr.strip().lower()
    
    # Temel geçerlilik kontrolleri
    if len(addr) < 5 or "@" not in addr:
        return None
    
    # Exchange internal adresleri filtrele
    invalid_tokens = ("/o=", "/ou=", "cn=", "/cn=", "x400:", "x500:")
    if any(token in addr for token in invalid_tokens):
        return None
    
    # Domain kontrolü
    if addr.count("@") != 1:
        return None
    
    local, domain = addr.split("@")
    if not local or not domain or "." not in domain:
        return None
    
    return addr


def resolve_smtp(address_entry, cache: Dict[str, Optional[str]]) -> Optional[str]:
    """
    AddressEntry'den SMTP adresini çözümle.
    Geliştirilmiş hata yakalama ve cache mekanizması.
    """
    if not address_entry:
        return None

    # Cache kontrolü
    cache_key = None
    with contextlib.suppress(AttributeError):
        cache_key = address_entry.ID
    
    if cache_key and cache_key in cache:
        return cache[cache_key]

    # Strateji 1: PropertyAccessor ile SMTP
    try:
        smtp = address_entry.PropertyAccessor.GetProperty(PROP_SMTP)
        if smtp:
            smtp = smtp.strip()
            if cache_key:
                cache[cache_key] = smtp
            return smtp
    except Exception as e:
        logging.debug("PropertyAccessor hatası: %s", e)

    # Strateji 2: Exchange User
    try:
        if address_entry.Type == "EX":
            ex_user = address_entry.GetExchangeUser()
            if ex_user and ex_user.PrimarySmtpAddress:
                smtp = ex_user.PrimarySmtpAddress.strip()
                if cache_key:
                    cache[cache_key] = smtp
                return smtp
    except Exception as e:
        logging.debug("Exchange User çözümleme hatası: %s", e)

    # Strateji 3: Direkt Address
    try:
        if address_entry.Address:
            smtp = address_entry.Address.strip()
            if "@" in smtp:  # Geçerli görünen adresler için
                if cache_key:
                    cache[cache_key] = smtp
                return smtp
    except Exception as e:
        logging.debug("Direkt Address hatası: %s", e)

    # Cache'e None kaydet (tekrar deneme yapma)
    if cache_key:
        cache[cache_key] = None
    
    return None


def extract_recipients(
    mail_item, 
    accumulator: Dict[str, str], 
    cache: Dict[str, Optional[str]],
    stats: ScanStats
) -> None:
    """Mail'den alıcıları çıkar ve accumulator'a ekle"""
    recipients = getattr(mail_item, "Recipients", None)
    if not recipients:
        return

    for recip in recipients:
        try:
            smtp = resolve_smtp(getattr(recip, "AddressEntry", None), cache)
            norm = normalize_email(smtp)
            
            if not norm or norm in accumulator:
                continue
            
            # İsim bilgisini al
            name = norm  # Default
            try:
                if hasattr(recip, "Name") and recip.Name:
                    name = recip.Name.strip()
            except Exception:
                pass
            
            accumulator[norm] = name
            
        except Exception as e:
            stats.errors_encountered += 1
            logging.debug("Recipient çözümleme hatası: %s", e)


def parse_address_field(value: Optional[str]) -> Set[str]:
    """String halindeki adres alanını parse et (To/Cc/Bcc için)"""
    if not value:
        return set()
    
    addresses = set()
    try:
        for _, addr in getaddresses([value]):
            norm = normalize_email(addr)
            if norm:
                addresses.add(norm)
    except Exception as e:
        logging.debug("Adres parse hatası: %s", e)
    
    return addresses


# ============================================================================
# TABLO MODU İŞLEME (GetTable - HIZLI)
# ============================================================================
def add_address_columns(table) -> None:
    """Tabloya gerekli adres kolonlarını ekle"""
    existing = {column.Name for column in table.Columns}
    for schema in ADDRESS_COLUMNS:
        if schema not in existing:
            try:
                table.Columns.Add(schema)
            except Exception as e:
                logging.debug("Kolon eklenemedi %s: %s", schema, e)


def iter_table_rows(table):
    """Tablo satırlarını iterate et"""
    while not table.EndOfTable:
        yield table.GetNextRow()


def process_table(
    items, 
    folder_name: str, 
    accumulator: Dict[str, str], 
    stats: ScanStats,
    max_items: Optional[int] = None
) -> int:
    """
    GetTable API ile hızlı tarama (sadece To/Cc/Bcc string alanları).
    İsim bilgisi olmayabilir, sadece e-posta adresleri.
    """
    try:
        table = items.GetTable()
        add_address_columns(table)
    except Exception as e:
        logging.warning("%s | GetTable başlatılamadı: %s", folder_name, e)
        return 0
    
    count = 0
    start_time = time.perf_counter()
    
    for row in iter_table_rows(table):
        count += 1
        
        # To/Cc/Bcc alanlarını parse et
        for schema in ADDRESS_COLUMNS:
            try:
                addresses = parse_address_field(row.get(schema))
                for addr in addresses:
                    if addr not in accumulator:
                        accumulator[addr] = addr  # İsim bilgisi yok, e-posta kendisi
            except Exception as e:
                stats.errors_encountered += 1
                logging.debug("Satır parse hatası: %s", e)
        
        # Progress log
        if count % PROGRESS_LOG_INTERVAL == 0:
            elapsed = time.perf_counter() - start_time
            rate = count / elapsed if elapsed > 0 else 0
            logging.info(
                "%s | %d mail işlendi (Table Mode, %.1f mail/sn)", 
                folder_name, count, rate
            )
        
        # Max items kontrolü
        if max_items and count >= max_items:
            logging.info("%s | Table mode %d kayıtta sınırlandı", folder_name, max_items)
            break
    
    elapsed = time.perf_counter() - start_time
    rate = count / elapsed if elapsed > 0 else 0
    logging.info(
        "%s | Table mode tamamlandı: %d mail, %.1f sn (%.1f mail/sn)", 
        folder_name, count, elapsed, rate
    )
    
    return count


# ============================================================================
# KLASİK ITERASYON (GetFirst/GetNext - YAVAS AMA EKSİKSİZ)
# ============================================================================
def process_iterative(
    items,
    folder_name: str,
    accumulator: Dict[str, str],
    cache: Dict[str, Optional[str]],
    stats: ScanStats,
    max_items: Optional[int] = None,
    checkpoint_callback = None
) -> int:
    """
    Klasik iterasyon ile tam recipient bilgisi çıkarma.
    Daha yavaş ama isim bilgisi dahil.
    """
    total = getattr(items, "Count", None)
    count = 0
    start_time = time.perf_counter()
    
    item = items.GetFirst()
    while item:
        try:
            # Sadece MailItem objelerini işle
            if getattr(item, "Class", None) == 43:
                extract_recipients(item, accumulator, cache, stats)
                count += 1
                
                # Checkpoint
                if checkpoint_callback and count % CHECKPOINT_INTERVAL == 0:
                    checkpoint_callback(accumulator, f"{folder_name}_checkpoint_{count}")
                
                # Progress log
                if count % PROGRESS_LOG_INTERVAL == 0:
                    elapsed = time.perf_counter() - start_time
                    rate = count / elapsed if elapsed > 0 else 0
                    
                    eta_text = ""
                    if total and count > 0 and rate > 0:
                        remaining = total - count
                        eta_seconds = remaining / rate
                        eta_text = f", ETA: ~{int(eta_seconds)}sn"
                    
                    logging.info(
                        "%s | %d/%s mail işlendi (%.1f mail/sn)%s",
                        folder_name, count, total or "?", rate, eta_text
                    )
                
                # Max items kontrolü
                if max_items and count >= max_items:
                    logging.info("%s | Klasik mode %d kayıtta sınırlandı", folder_name, max_items)
                    break
                    
        except Exception as exc:
            stats.errors_encountered += 1
            logging.debug("%s | Mail atlandı: %s", folder_name, exc)
        finally:
            try:
                item = items.GetNext()
            except Exception:
                break
            
            if max_items and count >= max_items:
                break
    
    elapsed = time.perf_counter() - start_time
    rate = count / elapsed if elapsed > 0 else 0
    logging.info(
        "%s | Klasik mode tamamlandı: %d mail, %.1f sn (%.1f mail/sn)",
        folder_name, count, elapsed, rate
    )
    
    return count


# ============================================================================
# KLASÖR İŞLEME
# ============================================================================
def process_folder(
    folder,
    accumulator: Dict[str, str],
    stats: ScanStats,
    restrict_days: Optional[int] = None,
    max_items: Optional[int] = None,
    use_table_mode: bool = True,
    checkpoint_callback = None
) -> None:
    """
    Bir klasörü tara. 
    use_table_mode=True ise önce GetTable dene (hızlı),
    sonra klasik iterasyon ile eksik kalan recipient'leri yakala.
    """
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    # Tarih filtresi
    if restrict_days is not None:
        try:
            cutoff = (datetime.now() - timedelta(days=restrict_days)).strftime("%m/%d/%Y %H:%M %p")
            items = items.Restrict(f"[ReceivedTime] >= '{cutoff}'")
            logging.info("%s | Son %d gün filtrelendi", folder.Name, restrict_days)
        except Exception as e:
            logging.warning("%s | Tarih filtresi uygulanamadı: %s", folder.Name, e)

    processed = 0
    cache: Dict[str, Optional[str]] = {}
    
    # Önce hızlı GetTable modu (sadece e-posta adresleri)
    if use_table_mode and hasattr(items, "GetTable"):
        logging.info("%s | Table mode başlatılıyor (hızlı tarama)", folder.Name)
        processed = process_table(items, folder.Name, accumulator, stats, max_items)
        stats.total_emails_scanned += processed
    
    # Sonra klasik iterasyon (isim bilgisi ve recipient detayları için)
    # Not: Eğer table mode tüm adresleri yakaladıysa, bu adım opsiyonel
    # Ancak isim bilgisi için klasik iterasyon şart
    logging.info("%s | Klasik mode başlatılıyor (detaylı tarama)", folder.Name)
    processed_classic = process_iterative(
        items, folder.Name, accumulator, cache, stats, max_items, checkpoint_callback
    )
    stats.total_emails_scanned += processed_classic
    stats.folders_scanned += 1


# ============================================================================
# CHECKPOINT VE EXPORT
# ============================================================================
def save_checkpoint(data: Dict[str, str], checkpoint_name: str, output_dir: Path) -> None:
    """Checkpoint CSV kaydet"""
    try:
        checkpoint_file = output_dir / f"checkpoint_{checkpoint_name}.csv"
        df = pd.DataFrame(sorted(data.items()), columns=["E-posta", "Ad"])
        df.to_csv(checkpoint_file, index=False, encoding='utf-8-sig')
        logging.debug("Checkpoint kaydedildi: %s", checkpoint_file)
    except Exception as e:
        logging.warning("Checkpoint kaydedilemedi: %s", e)


def export_to_excel(data: Dict[str, str], filename: Path, also_csv: bool = True) -> None:
    """Excel ve opsiyonel CSV export"""
    if not data:
        logging.warning("Kaydedilecek veri yok")
        return
    
    df = pd.DataFrame(sorted(data.items()), columns=["E-posta", "Ad"])
    filename.parent.mkdir(parents=True, exist_ok=True)
    
    # Excel export
    target = filename
    try:
        df.to_excel(target, index=False, engine="openpyxl")
        logging.info("✓ Excel kaydedildi: %s (%d adres)", target, len(df))
    except PermissionError:
        alt_name = filename.with_stem(f"{filename.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        df.to_excel(alt_name, index=False, engine="openpyxl")
        target = alt_name
        logging.warning("⚠ Dosya açıktı, alternatif kaydedildi: %s", alt_name)
    except Exception as e:
        logging.error("Excel kaydedilemedi: %s", e)
    
    # CSV export (büyük dosyalar için daha pratik)
    if also_csv:
        csv_target = filename.with_suffix('.csv')
        try:
            df.to_csv(csv_target, index=False, encoding='utf-8-sig')
            logging.info("✓ CSV kaydedildi: %s", csv_target)
        except Exception as e:
            logging.error("CSV kaydedilemedi: %s", e)


# ============================================================================
# ANA ÇALIŞTIRMA FONKSİYONU
# ============================================================================
def run(
    restrict_days: Optional[int] = None,
    output: Optional[Path] = None,
    verbose: bool = False,
    max_items: Optional[int] = None,
    use_table_mode: bool = True,
    save_checkpoints: bool = True,
    log_file: Optional[Path] = None,
    export_csv: bool = True
) -> None:
    """
    Outlook tarama ana fonksiyonu.
    
    Args:
        restrict_days: Kaç günlük mailler taransın (None=tümü)
        output: Çıktı dosya yolu
        verbose: Detaylı log
        max_items: Her klasörden max kaç mail (test için)
        use_table_mode: GetTable API kullan (hızlı)
        save_checkpoints: Checkpoint backup kaydet
        log_file: Log dosyası yolu
        export_csv: CSV de kaydet
    """
    configure_logging(verbose, log_file)
    logging.info("=" * 70)
    logging.info("OUTLOOK E-POSTA TARAYICI - GELİŞTİRİLMİŞ SÜRÜM")
    logging.info("=" * 70)

    # Outlook bağlantısı
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        logging.info("✓ Outlook bağlantısı kuruldu")
    except Exception as exc:
        logging.error("✗ Outlook açılamadı: %s", exc)
        return

    accumulator: Dict[str, str] = {}
    stats = ScanStats()
    
    # Output dizini
    if output is None:
        output = Path(os.getcwd()) / f"outlook_addresses_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    output_dir = output.parent
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Checkpoint callback
    def checkpoint_callback(data, name):
        if save_checkpoints:
            save_checkpoint(data, name, output_dir)
    
    # Store'ları tara
    stores = outlook.Stores
    logging.info("Toplam %d mail hesabı bulundu", stores.Count)
    
    for idx in range(1, stores.Count + 1):
        try:
            store = stores.Item(idx)
            store_name = store.DisplayName
            logging.info("-" * 70)
            logging.info("Mail Hesabı [%d/%d]: %s", idx, stores.Count, store_name)
            logging.info("-" * 70)

            # Gönderilmiş klasörü
            for folder_constant, tag in ((OL_FOLDER_SENT_MAIL, "Gönderilen"), (OL_FOLDER_ARCHIVE, "Arşiv")):
                try:
                    folder = store.GetDefaultFolder(folder_constant)
                    logging.info("📁 %s klasörü taranıyor...", tag)
                    process_folder(
                        folder, accumulator, stats, restrict_days, 
                        max_items, use_table_mode, checkpoint_callback
                    )
                except Exception as e:
                    logging.warning("⚠ %s klasörü alınamadı (%s): %s", tag, store_name, e)

            # Arşiv alias'ları (Arşiv, Archive vs.)
            try:
                root = store.GetRootFolder()
                for sub in root.Folders:
                    if sub.Name.strip().lower() in ARCHIVE_ALIASES:
                        logging.info("📁 %s (alias arşiv) taranıyor...", sub.Name)
                        process_folder(
                            sub, accumulator, stats, restrict_days,
                            max_items, use_table_mode, checkpoint_callback
                        )
            except Exception as e:
                logging.debug("Alias arşiv taraması hatası: %s", e)
                
        except Exception as e:
            logging.error("Store işleme hatası: %s", e)
            stats.errors_encountered += 1

    # İstatistikler
    stats.total_addresses_found = len(accumulator)
    logging.info("=" * 70)
    logging.info("TARAMA TAMAMLANDI")
    logging.info("=" * 70)
    logging.info(stats.summary())
    
    if not accumulator:
        logging.warning("⚠ Kaydedilecek e-posta adresi bulunamadı")
        return

    # Export
    export_to_excel(accumulator, output, also_csv=export_csv)
    
    logging.info("=" * 70)
    logging.info("✓ İşlem başarıyla tamamlandı!")
    logging.info("=" * 70)


# ============================================================================
# ÇALIŞTIRMA
# ============================================================================
if __name__ == "__main__":
    # Parametreleri buradan ayarlayabilirsin
    run(
        restrict_days=None,          # None = tüm mailler, 365 = son 1 yıl, vb.
        output=None,                 # None = otomatik isim, veya Path("cikti.xlsx")
        verbose=False,               # True = detaylı debug log
        max_items=None,              # Test için: 100, 1000, vb. | None = sınırsız
        use_table_mode=True,         # GetTable API kullan (hızlı)
        save_checkpoints=True,       # Her 1000 kayıtta checkpoint kaydet
        log_file=Path("outlook_scan.log"),  # Log dosyası (None = sadece konsol)
        export_csv=True              # Excel ile birlikte CSV de kaydet
    )

