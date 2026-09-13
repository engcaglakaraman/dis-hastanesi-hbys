import os
import random
import string
import io
import pyodbc
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from flask import Flask, render_template_string, request, redirect, url_for, flash, send_file

app = Flask(__name__)

app.secret_key = os.environ.get('SECRET_KEY') or os.urandom(24).hex()

CONN_STR = (
    r"DRIVER={ODBC Driver 17 for SQL Server};"
    r"SERVER=localhost;"
    r"DATABASE=DISHASTANESI;"
    r"Trusted_Connection=yes;"
)

def get_db():
    return pyodbc.connect(CONN_STR)

# ==========================================
# BASE HTML
# ==========================================
BASE_HTML = """
<!DOCTYPE html>
<html lang="tr" data-bs-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Diş Hastanesi Bilgi Portalı</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {
            --primary-navy: #0f172a;
            --accent-blue: #2563eb;
            --card-border: rgba(15, 23, 42, 0.08);
        }
        [data-bs-theme="dark"] {
            --primary-navy: #0b0f19;
            --card-border: rgba(255, 255, 255, 0.1);
        }
        body { 
            background-color: var(--bs-body-bg); 
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
            transition: background-color 0.3s ease, color 0.3s ease;
        }
        .navbar-custom {
            background: linear-gradient(135deg, #0a192f 0%, #1e3a8a 100%);
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
        }
        .card-custom {
            border: 1px solid var(--card-border);
            border-radius: 14px;
            box-shadow: 0 8px 16px rgba(0,0,0,0.03);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .card-custom:hover {
            transform: translateY(-2px);
            box-shadow: 0 12px 24px rgba(0,0,0,0.06);
        }

        .clickable-stat-card {
            cursor: pointer;
            text-decoration: none !important;
            transition: all 0.2s ease;
            display: block;
        }
        .clickable-stat-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 12px 25px rgba(0,0,0,0.2) !important;
            filter: brightness(1.05);
        }
        .card-active-border {
            border: 3px solid #facc15 !important;
            box-shadow: 0 0 18px rgba(250, 204, 21, 0.5) !important;
        }

        [data-bs-theme="dark"] .alert-danger {
            background-color: rgba(220, 38, 38, 0.2) !important;
            border: 1px solid #f87171 !important;
            color: #fca5a5 !important;
        }
        [data-bs-theme="dark"] .badge.bg-danger-subtle {
            background-color: rgba(239, 68, 68, 0.25) !important;
            color: #fca5a5 !important;
            border: 1px solid #ef4444 !important;
        }
        [data-bs-theme="dark"] .table-danger-subtle {
            background-color: rgba(220, 38, 38, 0.15) !important;
        }

        .floating-theme-switch {
            position: fixed;
            bottom: 24px;
            left: 24px;
            z-index: 9999;
            background: rgba(15, 23, 42, 0.85);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: 30px;
            padding: 6px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.25);
            cursor: pointer;
            user-select: none;
            transition: transform 0.2s ease, box-shadow 0.2s ease, background 0.3s ease;
        }
        .floating-theme-switch:hover {
            transform: translateY(-3px);
            box-shadow: 0 14px 30px rgba(0,0,0,0.35);
        }
        .switch-track {
            width: 32px;
            height: 60px;
            background: rgba(255,255,255,0.15);
            border-radius: 20px;
            position: relative;
            transition: background 0.3s ease;
        }
        .switch-thumb {
            width: 24px;
            height: 24px;
            background: #facc15;
            border-radius: 50%;
            position: absolute;
            top: 4px;
            left: 4px;
            transition: transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1), background 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 13px;
            color: #0f172a;
        }
        [data-bs-theme="dark"] .switch-thumb {
            transform: translateY(28px);
            background: #60a5fa;
            color: #ffffff;
        }
        [data-bs-theme="dark"] .floating-theme-switch {
            background: rgba(30, 41, 59, 0.9);
            border-color: rgba(255, 255, 255, 0.15);
        }
    </style>
</head>
<body class="d-flex flex-column min-vh-100">
    <div class="floating-theme-switch" id="themeToggle" title="Tema Modu">
        <div class="switch-track">
            <div class="switch-thumb" id="switchThumb">
                <i class="bi bi-sun-fill" id="thumbIcon"></i>
            </div>
        </div>
    </div>

    <nav class="navbar navbar-expand-xl navbar-dark navbar-custom mb-4 py-3">
        <div class="container-fluid px-4 px-lg-5 d-flex justify-content-between align-items-center">
            
            <a class="navbar-brand fw-bold d-flex align-items-center gap-2 me-4" href="/">
                <i class="bi bi-hospital fs-2 text-info"></i>
                <span class="fs-5">Diş Hastanesi <span class="fw-light opacity-75 fs-6">| Bilgi Portalı</span></span>
            </a>
            
            <form class="d-flex flex-grow-1 mx-lg-4" style="max-width: 480px;" action="/hastalar" method="GET">
                <div class="input-group">
                    <input class="form-control rounded-start-pill border-0 px-3 py-2" type="search" name="q" placeholder="Hasta No, Ad, TC veya Tıbbi Durum Ara..." aria-label="Search">
                    <button class="btn btn-info text-white rounded-end-pill px-3" type="submit">
                        <i class="bi bi-search"></i>
                    </button>
                </div>
            </form>

            <div class="d-flex align-items-center gap-3">
                <a href="/hastalar" class="btn btn-outline-light btn-sm rounded-pill px-3 py-2 fw-semibold">
                    <i class="bi bi-people me-1"></i> Tüm Hastalar
                </a>
                <a href="/kasa-raporlari" class="btn btn-outline-success btn-sm rounded-pill px-3 py-2 text-white fw-semibold">
                    <i class="bi bi-cash-stack me-1"></i> Kasa Raporları
                </a>
                <a href="/hasta-ekle" class="btn btn-outline-info btn-sm rounded-pill px-3 py-2 text-white fw-semibold">
                    <i class="bi bi-person-plus-fill me-1"></i> Yeni Hasta
                </a>
                <a href="/hekim-ekle" class="btn btn-outline-warning btn-sm rounded-pill px-3 py-2 text-white fw-semibold">
                    <i class="bi bi-person-badge-fill me-1"></i> Yeni Hekim
                </a>
                <a href="/stoklar" class="btn btn-outline-light btn-sm rounded-pill px-3 py-2 fw-semibold">
                    <i class="bi bi-box-seam me-1"></i> Malzeme & Stok
                </a>
                <a href="/randevu-ekle" class="btn btn-info text-white fw-bold px-4 py-2 shadow-sm rounded-pill">
                    <i class="bi bi-plus-circle me-1"></i> Randevu Al
                </a>
            </div>

        </div>
    </nav>

    <main class="container mb-5 flex-grow-1">
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, msg in messages %}
                    <div class="alert alert-{{ category }} alert-dismissible fade show rounded-3 shadow-sm" role="alert">
                        <i class="bi bi-info-circle-fill me-2"></i>{{ msg }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}
        {{ content | safe }}
    </main>

    <footer class="py-3 text-center text-muted border-top mt-auto small">
        <div class="container">
            © 2026 Diş Hastanesi Bilgi Yönetim Sistemi • SQL Server Entegre
        </div>
    </footer>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        const toggleBtn = document.getElementById('themeToggle');
        const thumbIcon = document.getElementById('thumbIcon');
        const html = document.documentElement;

        const savedTheme = localStorage.getItem('siteTheme') || 'light';
        setTheme(savedTheme);

        toggleBtn.addEventListener('click', () => {
            const newTheme = html.getAttribute('data-bs-theme') === 'light' ? 'dark' : 'light';
            setTheme(newTheme);
            localStorage.setItem('siteTheme', newTheme);
        });

        function setTheme(theme) {
            html.setAttribute('data-bs-theme', theme);
            if (theme === 'dark') {
                thumbIcon.className = 'bi bi-moon-stars-fill';
            } else {
                thumbIcon.className = 'bi bi-sun-fill';
            }
        }
    </script>
</body>
</html>
"""

# ==========================================
# INDEX
# ==========================================
INDEX_CONTENT = """
<div class="row g-3 mb-4">
    <div class="col-md-3">
        <div class="card card-custom p-4 bg-primary text-white border-0 shadow-sm" style="background: linear-gradient(135deg, #1e40af 0%, #1e3a8a 100%) !important;">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <p class="text-white-50 text-uppercase fw-bold small mb-1">Toplam Hasta</p>
                    <h2 class="display-6 fw-bold mb-0">{{ toplam_hasta }}</h2>
                </div>
                <i class="bi bi-people display-4 opacity-25"></i>
            </div>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card card-custom p-4 bg-info text-white border-0 shadow-sm" style="background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <p class="text-white-50 text-uppercase fw-bold small mb-1">Bugünkü Randevu</p>
                    <h2 class="display-6 fw-bold mb-0">{{ bugunku_randevu }}</h2>
                </div>
                <i class="bi bi-calendar-check display-4 opacity-25"></i>
            </div>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card card-custom p-4 bg-success text-white border-0 shadow-sm" style="background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <p class="text-white-50 text-uppercase fw-bold small mb-1">Uygulanan Tedavi</p>
                    <h2 class="display-6 fw-bold mb-0">{{ toplam_tedavi_sayisi }}</h2>
                </div>
                <i class="bi bi-capsule display-4 opacity-25"></i>
            </div>
        </div>
    </div>
    <div class="col-md-3">
        <div class="card card-custom p-4 bg-danger text-white border-0 shadow-sm" style="background: linear-gradient(135deg, #dc2626 0%, #991b1b 100%) !important;">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <p class="text-white-50 text-uppercase fw-bold small mb-1">Kritik Stok</p>
                    <h2 class="display-6 fw-bold mb-0">{{ kritik_stok_sayisi }}</h2>
                </div>
                <i class="bi bi-exclamation-triangle display-4 opacity-25"></i>
            </div>
        </div>
    </div>
</div>

<div class="card card-custom p-4">
    <div class="d-flex justify-content-between align-items-center mb-4">
        <div>
            <h4 class="fw-bold mb-0"><i class="bi bi-person-badge-fill text-primary me-2"></i>Aktif Hekim Kadrosu</h4>
            <small class="text-muted">Hastanede görev yapan poliklinik hekimleri</small>
        </div>
        <a href="/hekim-ekle" class="btn btn-outline-primary btn-sm rounded-pill"><i class="bi bi-plus-lg me-1"></i> Yeni Hekim Ekle</a>
    </div>
    <div class="row g-4">
        {% for hekim in hekimler %}
        <div class="col-md-6 col-lg-3">
            <div class="card h-100 card-custom border p-3">
                <div class="card-body p-0">
                    <div class="d-flex align-items-center gap-3 mb-3">
                        <div class="rounded-circle bg-primary-subtle text-primary p-3 d-flex align-items-center justify-content-center" style="width: 48px; height: 48px;">
                            <i class="bi bi-person-fill fs-4"></i>
                        </div>
                        <div>
                            <h6 class="card-title fw-bold mb-0">Dt. {{ hekim[1] }} {{ hekim[2] }}</h6>
                            <span class="badge bg-primary-subtle text-primary small mt-1">{{ hekim[3] }}</span>
                        </div>
                    </div>
                    <div class="small text-muted mb-2"><i class="bi bi-door-closed text-secondary me-1"></i> Oda: <strong>{{ hekim[4] }}</strong></div>
                    <div class="small text-muted mb-3"><i class="bi bi-telephone text-secondary me-1"></i> {{ hekim[5] }}</div>
                </div>
                <div class="card-footer bg-transparent border-0 p-0">
                    <a href="/hekim/{{ hekim[0] }}/randevular" class="btn btn-outline-primary btn-sm w-100 fw-semibold rounded-pill">
                        Randevu Listesi <i class="bi bi-arrow-right ms-1"></i>
                    </a>
                </div>
            </div>
        </div>
        {% endfor %}
    </div>
</div>
"""

# ==========================================
# İNTERAKTİF KASA RAPORLARI
# ==========================================
KASA_RAPORLARI_CONTENT = """
<div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-4">
    <div>
        <h4 class="fw-bold mb-1"><i class="bi bi-wallet2 text-success me-2"></i>Kasa ve Finansal Dönem Raporları</h4>
        <span class="text-muted small">Kartlara tıklayarak dönemi filtreleyebilir, Z-Raporu alabilir veya Excel indirebilirsiniz.</span>
    </div>
    <div class="d-flex gap-2">
        <a href="/kasa-raporlari/z-raporu" class="btn btn-dark text-white btn-sm rounded-pill px-3 fw-bold shadow-sm">
            <i class="bi bi-file-earmark-lock-fill me-1"></i> Gün Sonu (Z-Raporu)
        </a>
        <a href="/kasa-raporlari/excel?filtre={{ aktif_filtre }}&baslangic={{ baslangic }}&bitis={{ bitis }}" class="btn btn-success text-white btn-sm rounded-pill px-3 fw-bold shadow-sm">
            <i class="bi bi-file-earmark-excel-fill me-1"></i> Excel (.xlsx)
        </a>
        <a href="/kasa-raporlari/hekim-performans" class="btn btn-warning text-dark btn-sm rounded-pill px-3 fw-bold shadow-sm">
            <i class="bi bi-graph-up-arrow me-1"></i> Hekim Performans
        </a>
        <a href="/" class="btn btn-outline-secondary btn-sm rounded-pill"><i class="bi bi-arrow-left me-1"></i> Ana Sayfa</a>
    </div>
</div>

<div class="row g-3 mb-4">
    <div class="col-md-4">
        <a href="/kasa-raporlari?filtre=gunluk" class="clickable-stat-card card card-custom p-4 bg-primary text-white border-0 shadow-sm {% if aktif_filtre == 'gunluk' %}card-active-border{% endif %}" style="background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <p class="text-white-50 text-uppercase fw-bold small mb-1">Bugünkü Tahsilat {% if aktif_filtre == 'gunluk' %}<i class="bi bi-check-circle-fill text-warning ms-1"></i>{% endif %}</p>
                    <h2 class="display-6 fw-bold mb-0">{{ "%.2f"|format(gunluk_tutar) }} ₺</h2>
                    <small class="text-white-50 mt-2 d-block">{{ gunluk_adet }} İşlem • Filtrele</small>
                </div>
                <i class="bi bi-calendar-check display-4 opacity-25"></i>
            </div>
        </a>
    </div>
    <div class="col-md-4">
        <a href="/kasa-raporlari?filtre=haftalik" class="clickable-stat-card card card-custom p-4 bg-info text-white border-0 shadow-sm {% if aktif_filtre == 'haftalik' %}card-active-border{% endif %}" style="background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <p class="text-white-50 text-uppercase fw-bold small mb-1">Bu Haftaki Ciro (Son 7 Gün) {% if aktif_filtre == 'haftalik' %}<i class="bi bi-check-circle-fill text-warning ms-1"></i>{% endif %}</p>
                    <h2 class="display-6 fw-bold mb-0">{{ "%.2f"|format(haftalik_tutar) }} ₺</h2>
                    <small class="text-white-50 mt-2 d-block">{{ haftalik_adet }} İşlem • Filtrele</small>
                </div>
                <i class="bi bi-calendar-week display-4 opacity-25"></i>
            </div>
        </a>
    </div>
    <div class="col-md-4">
        <a href="/kasa-raporlari?filtre=aylik" class="clickable-stat-card card card-custom p-4 bg-success text-white border-0 shadow-sm {% if aktif_filtre == 'aylik' %}card-active-border{% endif %}" style="background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <p class="text-white-50 text-uppercase fw-bold small mb-1">Bu Ayki Toplam Ciro (Son 30 Gün) {% if aktif_filtre == 'aylik' %}<i class="bi bi-check-circle-fill text-warning ms-1"></i>{% endif %}</p>
                    <h2 class="display-6 fw-bold mb-0">{{ "%.2f"|format(aylik_tutar) }} ₺</h2>
                    <small class="text-white-50 mt-2 d-block">{{ aylik_adet }} İşlem • Filtrele</small>
                </div>
                <i class="bi bi-calendar-month display-4 opacity-25"></i>
            </div>
        </a>
    </div>
</div>

<div class="row g-4 mb-4">
    <div class="col-lg-5">
        <div class="card card-custom p-4 h-100 d-flex flex-column justify-content-center align-items-center">
            <h6 class="fw-bold text-secondary mb-3"><i class="bi bi-pie-chart-fill me-1"></i> Ödeme Yöntemleri Dağılımı</h6>
            <div style="width: 220px; height: 220px;">
                <canvas id="odemeGrafik"></canvas>
            </div>
        </div>
    </div>
    <div class="col-lg-7">
        <div class="card card-custom p-4 h-100 d-flex flex-column justify-content-center">
            <h6 class="fw-bold text-secondary mb-3"><i class="bi bi-funnel-fill me-1"></i> Özel Tarih Aralığı Filtresi</h6>
            <form action="/kasa-raporlari" method="GET" class="row g-3">
                <input type="hidden" name="filtre" value="ozel">
                <div class="col-md-6">
                    <label class="form-label small fw-semibold">Başlangıç Tarihi</label>
                    <input type="date" name="baslangic" class="form-control" value="{{ baslangic or '' }}" required>
                </div>
                <div class="col-md-6">
                    <label class="form-label small fw-semibold">Bitiş Tarihi</label>
                    <input type="date" name="bitis" class="form-control" value="{{ bitis or '' }}" required>
                </div>
                <div class="col-12 mt-3">
                    <button type="submit" class="btn btn-primary rounded-pill px-4 fw-bold w-100 shadow-sm">
                        <i class="bi bi-search me-1"></i> Raporu Getir
                    </button>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
    const ctx = document.getElementById('odemeGrafik').getContext('2d');
    const odemeGrafik = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: {{ grafik_etiketleri | safe }},
            datasets: [{
                data: {{ grafik_verileri | safe }},
                backgroundColor: ['#2563eb', '#059669', '#f59e0b', '#dc2626', '#7c3aed'],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { boxWidth: 12, font: { size: 11 } }
                }
            }
        }
    });
</script>

<div class="card card-custom p-4 mb-4">
    <div class="d-flex justify-content-between align-items-center mb-3">
        <div>
            <h5 class="fw-bold mb-0"><i class="bi bi-clock-history text-primary me-2"></i>Tahsilat Belgeleri & Dökümü</h5>
            <span class="badge bg-secondary-subtle text-body mt-1">Gösterilen Filtre: <strong>{{ filtre_aciklama }}</strong> (Toplam: <strong>{{ "%.2f"|format(filtrelenen_toplam) }} ₺</strong>)</span>
        </div>
    </div>
    <div class="table-responsive">
        <table class="table table-hover align-middle">
            <thead class="table-light">
                <tr>
                    <th>Belge / İşlem No</th>
                    <th>Tarih</th>
                    <th>Hasta No & Adı</th>
                    <th>Ödeme Yöntemi</th>
                    <th>Açıklama</th>
                    <th class="text-end">Tutar</th>
                    <th class="text-end">İşlem</th>
                </tr>
            </thead>
            <tbody>
                {% for o in odemeler %}
                <tr>
                    <td><span class="badge bg-secondary-subtle text-body">MAK-{{ "%05d"|format(o[0]) }}</span></td>
                    <td class="fw-semibold">{{ o[1] }}</td>
                    <td>
                        <span class="badge bg-primary-subtle text-primary fw-bold me-1">{{ o[6] or ('HST-' ~ o[5]) }}</span>
                        <strong>{{ o[7] }} {{ o[8] }}</strong>
                    </td>
                    <td><span class="badge bg-light text-secondary border">{{ o[2] }}</span></td>
                    <td class="small text-muted">{{ o[4] or 'Genel Tahsilat' }}</td>
                    <td class="text-end fw-bold text-success">{{ "%.2f"|format(o[3]) }} ₺</td>
                    <td class="text-end">
                        <a href="/odeme-makbuz/{{ o[0] }}" target="_blank" class="btn btn-outline-primary btn-sm rounded-pill px-3">
                            <i class="bi bi-printer me-1"></i> Yazdır
                        </a>
                    </td>
                </tr>
                {% else %}
                <tr>
                    <td colspan="7" class="text-center text-muted py-4">Seçilen döneme ait tahsilat kaydı bulunamadı.</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>
"""

# ==========================================
# Z-RAPORU
# ==========================================
Z_RAPORU_HTML = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <title>Gün Sonu Z-Raporu - Kasa Mutabakatı</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { font-family: 'Inter', system-ui, sans-serif; color: #1e293b; }
        .z-card { max-width: 650px; margin: 40px auto; border: 2px solid #0f172a; border-radius: 12px; padding: 40px; }
        @media print {
            .no-print { display: none !important; }
            .z-card { border: none; margin: 0 auto; box-shadow: none; padding: 0; }
        }
    </style>
</head>
<body class="bg-light">
    <div class="container">
        <div class="z-card bg-white shadow-sm">
            <div class="text-center border-bottom pb-3 mb-4">
                <h4 class="fw-bold mb-1 text-dark">DİŞ HASTANESİ</h4>
                <span class="text-muted small">Günlük Vezne & Kasa Kapatma (Z-Raporu)</span>
                <div class="mt-2"><span class="badge bg-dark">Tarih: {{ bugun }}</span></div>
            </div>

            <div class="row g-3 mb-4">
                <div class="col-6">
                    <span class="text-muted small d-block">Toplam İşlem Adedi:</span>
                    <strong class="fs-5">{{ toplam_islem }} Adet</strong>
                </div>
                <div class="col-6 text-end">
                    <span class="text-muted small d-block">Günlük Toplam Ciro:</span>
                    <strong class="fs-5 text-success">{{ "%.2f"|format(toplam_ciro) }} ₺</strong>
                </div>
            </div>

            <h6 class="fw-bold border-bottom pb-2 mb-3">Ödeme Yöntemine Göre Dağılım</h6>
            <table class="table table-bordered mb-4 small">
                <thead class="table-light">
                    <tr>
                        <th>Ödeme Yöntemi</th>
                        <th class="text-center">İşlem Sayısı</th>
                        <th class="text-end">Toplam Tutar</th>
                    </tr>
                </thead>
                <tbody>
                    {% for d in detaylar %}
                    <tr>
                        <td class="fw-bold">{{ d[0] }}</td>
                        <td class="text-center">{{ d[1] }}</td>
                        <td class="text-end fw-semibold text-success">{{ "%.2f"|format(d[2]) }} ₺</td>
                    </tr>
                    {% else %}
                    <tr>
                        <td colspan="3" class="text-center text-muted">Bugün henüz tahsilat yapılmadı.</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>

            <div class="row mt-5 pt-4 border-top text-center small text-muted">
                <div class="col-6">
                    <span>Veznedar / Sorumlu</span><br><br><br>
                    <strong>İmza</strong>
                </div>
                <div class="col-6">
                    <span>Mesul Müdür / Yönetim</span><br><br><br>
                    <strong>İmza</strong>
                </div>
            </div>

            <div class="text-center mt-4 no-print">
                <button onclick="window.print()" class="btn btn-dark rounded-pill px-4 fw-bold">
                    <i class="bi bi-printer me-1"></i> Z-Raporunu Yazdır
                </button>
                <a href="/kasa-raporlari" class="btn btn-outline-secondary rounded-pill px-3 ms-2">
                    Kasa Raporlarına Dön
                </a>
            </div>
        </div>
    </div>
</body>
</html>
"""

# ==========================================
# HEKİM PERFORMANS
# ==========================================
HEKIM_PERFORMANS_CONTENT = """
<div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-4">
    <div>
        <h4 class="fw-bold mb-1"><i class="bi bi-graph-up-arrow text-warning me-2"></i>Hekim Bazlı Ciro & Performans Analizi</h4>
        <span class="text-muted small">Poliklinik hekim kadrosunun baktığı hasta sayıları, uyguladığı tedaviler ve ciro katkıları</span>
    </div>
    <div class="d-flex gap-2">
        <a href="/kasa-raporlari" class="btn btn-outline-success btn-sm rounded-pill px-3 fw-bold">
            <i class="bi bi-cash-stack me-1"></i> Kasa Raporlarına Dön
        </a>
        <a href="/" class="btn btn-outline-secondary btn-sm rounded-pill"><i class="bi bi-arrow-left me-1"></i> Ana Sayfa</a>
    </div>
</div>

<div class="row g-3 mb-4">
    <div class="col-md-4">
        <div class="card card-custom p-4 bg-primary text-white border-0 shadow-sm" style="background: linear-gradient(135deg, #1e40af 0%, #1e3a8a 100%) !important;">
            <span class="text-white-50 text-uppercase fw-bold small">Kadro Büyüklüğü</span>
            <h3 class="fw-bold mb-0 mt-1">{{ hekim_performans|length }} Aktif Hekim</h3>
            <small class="text-white-50 mt-1 d-block">Poliklinik Hekim Kadrosu</small>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card card-custom p-4 bg-success text-white border-0 shadow-sm" style="background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;">
            <span class="text-white-50 text-uppercase fw-bold small">Toplam Üretilen Tedavi Cirosu</span>
            <h3 class="fw-bold mb-0 mt-1">{{ "%.2f"|format(toplam_ciro) }} ₺</h3>
            <small class="text-white-50 mt-1 d-block">Tüm Hekimlerin Toplam Geliri</small>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card card-custom p-4 bg-warning text-dark border-0 shadow-sm" style="background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%) !important;">
            <span class="text-dark-50 text-uppercase fw-bold small">Lider Hekim</span>
            <h4 class="fw-bold mb-0 mt-1">{{ lider_hekim or 'Henüz Veri Yok' }}</h4>
            <small class="text-dark-50 mt-1 d-block">En Yüksek Ciro Üreten</small>
        </div>
    </div>
</div>

<div class="card card-custom p-4">
    <h5 class="fw-bold mb-3"><i class="bi bi-trophy-fill text-warning me-2"></i>Hekim Performans Sıralaması ve Dağılımı</h5>
    <div class="table-responsive">
        <table class="table table-hover align-middle">
            <thead class="table-light">
                <tr>
                    <th>#</th>
                    <th>Diş Hekimi</th>
                    <th>Branş</th>
                    <th>Baktığı Hasta</th>
                    <th>Tedavi Adedi</th>
                    <th style="min-width: 200px;">Ciro Katkı Oranı</th>
                    <th class="text-end">Üretilen Toplam Tutar</th>
                </tr>
            </thead>
            <tbody>
                {% for hp in hekim_performans %}
                <tr>
                    <td><span class="badge bg-secondary-subtle text-body">{{ loop.index }}</span></td>
                    <td>
                        <strong class="fs-6">Dt. {{ hp[1] }} {{ hp[2] }}</strong>
                        <div class="small text-muted">Oda: {{ hp[4] }}</div>
                    </td>
                    <td><span class="badge bg-primary-subtle text-primary border border-primary-subtle">{{ hp[3] }}</span></td>
                    <td class="fw-bold">{{ hp[5] }} Hasta</td>
                    <td class="fw-semibold">{{ hp[6] }} İşlem</td>
                    <td>
                        <div class="d-flex justify-content-between small fw-bold mb-1">
                            <span>Katkı: %{{ hp[8] }}</span>
                        </div>
                        <div class="progress" style="height: 8px;">
                            <div class="progress-bar {% if loop.index == 1 %}bg-warning{% elif loop.index == 2 %}bg-info{% else %}bg-primary{% endif %}" role="progressbar" style="width: {{ hp[8] }}%;"></div>
                        </div>
                    </td>
                    <td class="text-end fw-bold text-success fs-6">{{ "%.2f"|format(hp[7]) }} ₺</td>
                </tr>
                {% else %}
                <tr>
                    <td colspan="7" class="text-center text-muted py-4">Kayıtlı performans verisi bulunamadı.</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>
"""

MAKBUZ_HTML = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <title>Tahsilat Makbuzu #MAK-{{ "%05d"|format(odeme[0]) }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { font-family: 'Inter', system-ui, sans-serif; color: #1e293b; }
        .receipt-card { max-width: 650px; margin: 40px auto; border: 2px dashed #94a3b8; border-radius: 12px; padding: 30px; }
        @media print {
            .no-print { display: none !important; }
            .receipt-card { border: 1px solid #333; margin: 0 auto; box-shadow: none; }
        }
    </style>
</head>
<body class="bg-light">
    <div class="container">
        <div class="receipt-card bg-white shadow-sm">
            <div class="d-flex justify-content-between align-items-center border-bottom pb-3 mb-3">
                <div>
                    <h4 class="fw-bold mb-0 text-primary">DİŞ HASTANESİ</h4>
                    <span class="small text-muted">Ağız ve Diş Sağlığı Bilgi Yönetim Sistemi</span>
                </div>
                <div class="text-end">
                    <h5 class="fw-bold mb-0">TAHSİLAT MAKBUZU</h5>
                    <span class="badge bg-secondary">No: MAK-{{ "%05d"|format(odeme[0]) }}</span>
                </div>
            </div>

            <div class="row g-3 mb-4 small">
                <div class="col-6">
                    <span class="text-muted d-block">Hasta Adı Soyadı:</span>
                    <strong class="fs-6">{{ odeme[5] }} {{ odeme[6] }}</strong>
                </div>
                <div class="col-6 text-end">
                    <span class="text-muted d-block">Hasta No:</span>
                    <strong>{{ odeme[7] or ('HST-' ~ odeme[4]) }}</strong>
                </div>
                <div class="col-6">
                    <span class="text-muted d-block">Tahsilat Tarihi:</span>
                    <strong>{{ odeme[1] }}</strong>
                </div>
                <div class="col-6 text-end">
                    <span class="text-muted d-block">Ödeme Yöntemi:</span>
                    <strong>{{ odeme[2] }}</strong>
                </div>
            </div>

            <div class="card bg-light border-0 p-3 mb-4">
                <div class="d-flex justify-content-between align-items-center">
                    <span class="fw-bold">TAHSİL EDİLEN TUTAR:</span>
                    <h3 class="fw-bold text-success mb-0">{{ "%.2f"|format(odeme[3]) }} ₺</h3>
                </div>
                <div class="mt-2 small text-muted">
                    <span>Açıklama: </span><strong>{{ odeme[8] or 'Klinik Tedavi ve İşlem Tahsilatı' }}</strong>
                </div>
            </div>

            <div class="row mt-4 pt-4 border-top text-center small text-muted">
                <div class="col-6">
                    <span>Tahsil Eden Yetkili</span><br><br>
                    <strong>Vezne / Muhasebe</strong>
                </div>
                <div class="col-6">
                    <span>Hasta / Ödeyen</span><br><br>
                    <strong>İmza</strong>
                </div>
            </div>

            <div class="mt-4 pt-3 border-top text-center text-muted" style="font-size: 11px;">
                Bu belge elektronik ortamda kayıt altına alınmış resmi tahsilat dökümüdür.
            </div>

            <div class="text-center mt-4 no-print">
                <button onclick="window.print()" class="btn btn-primary btn-sm rounded-pill px-4 fw-bold">
                    <i class="bi bi-printer me-1"></i> Yazdır / PDF İndir
                </button>
                <button onclick="window.close()" class="btn btn-outline-secondary btn-sm rounded-pill px-3">
                    Kapat
                </button>
            </div>
        </div>
    </div>
</body>
</html>
"""

PLAN_TEKLIF_HTML = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <title>Tedavi Planı & Fiyat Teklifi #PLN-{{ "%04d"|format(plan[0]) }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { font-family: 'Inter', system-ui, sans-serif; color: #1e293b; }
        .quote-card { max-width: 750px; margin: 30px auto; border: 1px solid #cbd5e1; border-radius: 14px; padding: 40px; }
        @media print {
            .no-print { display: none !important; }
            .quote-card { border: none; margin: 0 auto; box-shadow: none; padding: 0; }
        }
    </style>
</head>
<body class="bg-light">
    <div class="container">
        <div class="quote-card bg-white shadow-sm">
            <div class="d-flex justify-content-between align-items-center border-bottom pb-4 mb-4">
                <div>
                    <h3 class="fw-bold text-primary mb-1">DİŞ HASTANESİ</h3>
                    <span class="text-muted small">Klinik Tedavi Planı, Seans Çizelgesi & Fiyat Teklifi</span>
                </div>
                <div class="text-end">
                    <span class="badge bg-primary fs-6 px-3 py-2">PLN-{{ "%04d"|format(plan[0]) }}</span>
                    <div class="small text-muted mt-1">Tarih: {{ plan[8] }}</div>
                </div>
            </div>

            <div class="row g-3 mb-4 p-3 bg-light rounded-3">
                <div class="col-6">
                    <span class="text-muted small d-block">Hasta Adı Soyadı:</span>
                    <strong class="fs-5">{{ plan[9] }} {{ plan[10] }}</strong>
                    <div class="small text-muted">Protokol No: {{ plan[11] or ('HST-' ~ plan[1]) }}</div>
                </div>
                <div class="col-6 text-end">
                    <span class="text-muted small d-block">Sorumlu Hekim:</span>
                    <strong class="fs-6">Dt. {{ plan[12] }} {{ plan[13] }}</strong>
                </div>
            </div>

            <div class="mb-4">
                <h5 class="fw-bold border-bottom pb-2 mb-3">Tedavi Planı Detayı</h5>
                <h6 class="fw-bold text-primary">{{ plan[3] }}</h6>
                <p class="text-secondary small">{{ plan[4] or 'Özel plan açıklaması girilmedi.' }}</p>
            </div>

            <table class="table table-bordered mb-4">
                <thead class="table-light">
                    <tr>
                        <th>Planlanan Seans Sayısı</th>
                        <th>Tamamlanan Seans</th>
                        <th>Plan Durumu</th>
                        <th class="text-end">Tahmini Toplam Tutar</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td class="fw-bold">{{ plan[5] }} Seans</td>
                        <td>{{ plan[6] }} Seans</td>
                        <td><span class="badge bg-secondary">{{ plan[7] }}</span></td>
                        <td class="text-end fw-bold fs-5 text-primary">{{ "%.2f"|format(plan[2]) }} ₺</td>
                    </tr>
                </tbody>
            </table>

            <div class="alert alert-info py-2 small mb-4">
                <i class="bi bi-info-circle me-1"></i> Bu tedavi planı ve teklifi klinik muayene doğrultusunda hazırlanmış olup tahmini seans sürelerini içerir.
            </div>

            <div class="row mt-5 pt-4 border-top text-center small text-muted">
                <div class="col-6">
                    <span>Hazırlayan Diş Hekimi</span><br><br><br>
                    <strong>Dt. {{ plan[12] }} {{ plan[13] }}</strong>
                </div>
                <div class="col-6">
                    <span>Tedavi Planını Onaylayan Hasta</span><br><br><br>
                    <strong>{{ plan[9] }} {{ plan[10] }}</strong>
                </div>
            </div>

            <div class="text-center mt-4 no-print">
                <button onclick="window.print()" class="btn btn-primary rounded-pill px-4 fw-bold">
                    <i class="bi bi-printer me-1"></i> Teklifi Yazdır / PDF
                </button>
                <button onclick="window.close()" class="btn btn-outline-secondary rounded-pill px-3 ms-2">
                    Kapat
                </button>
            </div>
        </div>
    </div>
</body>
</html>
"""

HASTALAR_CONTENT = """
<div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-4">
    <div>
        <h4 class="fw-bold mb-1"><i class="bi bi-people-fill text-primary me-2"></i>Kayıtlı Tüm Hastalar</h4>
        <span class="text-muted small">{% if arama %}Arama Sonucu: "<strong>{{ arama }}</strong>"{% else %}Sistemde kayıtlı hasta listesi{% endif %}</span>
    </div>
    <div class="d-flex gap-2">
        <form class="d-flex" action="/hastalar" method="GET">
            <input class="form-control form-control-sm rounded-start-pill px-3" type="search" name="q" value="{{ arama or '' }}" placeholder="Hasta No, Ad, TC Ara...">
            <button class="btn btn-primary btn-sm rounded-end-pill px-3" type="submit"><i class="bi bi-search"></i></button>
        </form>
        <a href="/hasta-ekle" class="btn btn-primary btn-sm rounded-pill px-3"><i class="bi bi-person-plus me-1"></i> Yeni Hasta</a>
        <a href="/" class="btn btn-outline-secondary btn-sm rounded-pill"><i class="bi bi-arrow-left me-1"></i> Ana Sayfa</a>
    </div>
</div>

<div class="card card-custom p-4">
    <div class="table-responsive">
        <table class="table table-hover align-middle">
            <thead class="table-light">
                <tr>
                    <th>Hasta No</th>
                    <th>T.C. Kimlik No</th>
                    <th>Adı Soyadı</th>
                    <th>Tıbbi Durum / Alerji</th>
                    <th>Doğum Tarihi</th>
                    <th>Telefon</th>
                    <th class="text-end">İşlemler</th>
                </tr>
            </thead>
            <tbody>
                {% for h in hastalar %}
                <tr>
                    <td><span class="badge bg-primary-subtle text-primary border border-primary-subtle fw-bold">{{ h[9] or ('HST-' ~ h[0]) }}</span></td>
                    <td class="fw-bold">{{ h[1] or '-' }}</td>
                    <td class="fw-semibold">{{ h[2] }} {{ h[3] }}</td>
                    <td>
                        {% if h[8] %}
                            <span class="badge bg-danger-subtle text-danger border border-danger-subtle px-2 py-1">
                                <i class="bi bi-exclamation-octagon-fill me-1"></i>{{ h[8] }}
                            </span>
                        {% else %}
                            <span class="badge bg-light text-muted border">Belirtilmedi</span>
                        {% endif %}
                    </td>
                    <td>{{ h[4] or '-' }}</td>
                    <td>{{ h[6] or '-' }}</td>
                    <td class="text-end">
                        <a href="/hasta/{{ h[0] }}/gecmis" class="btn btn-outline-info btn-sm rounded-pill px-2 me-1" title="Tüm Tıbbi Geçmiş">
                            <i class="bi bi-clock-history me-1"></i> Geçmiş
                        </a>
                        <a href="/hasta/{{ h[0] }}" class="btn btn-primary btn-sm rounded-pill px-3 me-1">
                            <i class="bi bi-file-medical me-1"></i> Dosya
                        </a>
                        <a href="/hasta/{{ h[0] }}/duzenle" class="btn btn-outline-warning btn-sm rounded-pill px-2 me-1 text-dark" title="Bilgileri Düzenle">
                            <i class="bi bi-pencil-square"></i>
                        </a>
                        <a href="/hasta-sil/{{ h[0] }}" class="btn btn-outline-danger btn-sm rounded-pill px-2" onclick="return confirm('Bu hastayı ve tüm tıbbi geçmişini kalıcı olarak silmek istiyor musunuz?');">
                            <i class="bi bi-trash"></i>
                        </a>
                    </td>
                </tr>
                {% else %}
                <tr>
                    <td colspan="7" class="text-center text-muted py-4">Aradığınız kriterde hasta bulunamadı.</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>
"""

HASTA_EKLE_CONTENT = """
<div class="row justify-content-center">
    <div class="col-md-8 col-lg-7">
        <div class="card card-custom p-4 p-md-5">
            <div class="text-center mb-4">
                <div class="rounded-circle bg-primary-subtle text-primary p-3 d-inline-flex mb-2">
                    <i class="bi bi-person-plus-fill fs-3"></i>
                </div>
                <h4 class="fw-bold">Yeni Hasta Kayıt Formu</h4>
                <p class="text-muted small">Ad ve Soyad haricindeki tüm alanlar isteğe bağlıdır</p>
            </div>
            
            <form method="POST">
                <div class="row g-3 mb-3">
                    <div class="col-md-6">
                        <label class="form-label fw-semibold text-primary"><i class="bi bi-tag-fill me-1"></i>Hasta Numarası</label>
                        <input type="text" name="hasta_no" class="form-control border-primary-subtle fw-bold" placeholder="Örn: HST-1001 (İsteğe bağlı)">
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">T.C. Kimlik No</label>
                        <input type="text" name="tc_no" class="form-control" maxlength="11" placeholder="11 Haneli TC (İsteğe bağlı)">
                    </div>
                </div>

                <div class="row g-3 mb-3">
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">Adı <span class="text-danger">*</span></label>
                        <input type="text" name="ad" class="form-control" placeholder="Adı" required>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">Soyadı <span class="text-danger">*</span></label>
                        <input type="text" name="soyad" class="form-control" placeholder="Soyadı" required>
                    </div>
                </div>

                <div class="mb-3">
                    <label class="form-label fw-semibold text-danger"><i class="bi bi-heart-pulse-fill me-1"></i>Tıbbi Durum / Kronik Rahatsızlık & Alerji Notu</label>
                    <input type="text" name="tibbi_not" class="form-control border-danger-subtle" placeholder="Örn: Diyabet (Tip 2), Penisilin Alerjisi...">
                </div>

                <div class="row g-3 mb-3">
                    <div class="col-md-4">
                        <label class="form-label fw-semibold">Doğum Tarihi</label>
                        <input type="date" name="dogum_tarihi" class="form-control">
                    </div>
                    <div class="col-md-4">
                        <label class="form-label fw-semibold">Cinsiyet</label>
                        <select name="cinsiyet" class="form-select">
                            <option value="Belirtilmemiş">Belirtilmemiş</option>
                            <option value="Kadın">Kadın</option>
                            <option value="Erkek">Erkek</option>
                        </select>
                    </div>
                    <div class="col-md-4">
                        <label class="form-label fw-semibold">Telefon Numarası</label>
                        <input type="text" name="telefon" class="form-control" placeholder="05XXXXXXXXX">
                    </div>
                </div>

                <div class="mb-3">
                    <label class="form-label fw-semibold">E-Posta Adresi</label>
                    <input type="email" name="email" class="form-control" placeholder="ornek@mail.com">
                </div>

                <div class="mb-4">
                    <label class="form-label fw-semibold">İkametgah Adresi</label>
                    <textarea name="adres" class="form-control" rows="2" placeholder="Açık adres..."></textarea>
                </div>

                <div class="d-flex gap-2">
                    <a href="/hastalar" class="btn btn-outline-secondary w-50 rounded-pill">İptal</a>
                    <button type="submit" class="btn btn-primary w-50 fw-bold rounded-pill shadow-sm">Hastayı Sisteme Kaydet</button>
                </div>
            </form>
        </div>
    </div>
</div>
"""

HASTA_DUZENLE_CONTENT = """
<div class="row justify-content-center">
    <div class="col-md-8 col-lg-7">
        <div class="card card-custom p-4 p-md-5">
            <div class="text-center mb-4">
                <div class="rounded-circle bg-warning-subtle text-warning-emphasis p-3 d-inline-flex mb-2">
                    <i class="bi bi-pencil-square fs-3"></i>
                </div>
                <h4 class="fw-bold">Hasta Bilgilerini Güncelle</h4>
                <p class="text-muted small">{{ hasta[2] }} {{ hasta[3] }} isimli hastanın kayıt düzenleme paneli</p>
            </div>
            
            <form method="POST">
                <div class="row g-3 mb-3">
                    <div class="col-md-6">
                        <label class="form-label fw-semibold text-primary"><i class="bi bi-tag-fill me-1"></i>Hasta Numarası</label>
                        <input type="text" name="hasta_no" class="form-control border-primary-subtle fw-bold" value="{{ hasta[9] or '' }}">
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">T.C. Kimlik No</label>
                        <input type="text" name="tc_no" class="form-control" maxlength="11" value="{{ hasta[1] or '' }}">
                    </div>
                </div>

                <div class="row g-3 mb-3">
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">Adı <span class="text-danger">*</span></label>
                        <input type="text" name="ad" class="form-control" value="{{ hasta[2] }}" required>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">Soyadı <span class="text-danger">*</span></label>
                        <input type="text" name="soyad" class="form-control" value="{{ hasta[3] }}" required>
                    </div>
                </div>

                <div class="mb-3">
                    <label class="form-label fw-semibold text-danger"><i class="bi bi-heart-pulse-fill me-1"></i>Tıbbi Durum / Alerji / Diyabet Notu</label>
                    <input type="text" name="tibbi_not" class="form-control border-danger-subtle" value="{{ hasta[8] or '' }}" placeholder="Örn: Diyabet, Penisilin Alerjisi...">
                </div>

                <div class="row g-3 mb-3">
                    <div class="col-md-4">
                        <label class="form-label fw-semibold">Doğum Tarihi</label>
                        <input type="date" name="dogum_tarihi" class="form-control" value="{{ hasta[4] or '' }}">
                    </div>
                    <div class="col-md-4">
                        <label class="form-label fw-semibold">Cinsiyet</label>
                        <select name="cinsiyet" class="form-select">
                            <option value="Belirtilmemiş" {% if hasta[5] == 'Belirtilmemiş' %}selected{% endif %}>Belirtilmemiş</option>
                            <option value="Kadın" {% if hasta[5] == 'Kadın' %}selected{% endif %}>Kadın</option>
                            <option value="Erkek" {% if hasta[5] == 'Erkek' %}selected{% endif %}>Erkek</option>
                        </select>
                    </div>
                    <div class="col-md-4">
                        <label class="form-label fw-semibold">Telefon</label>
                        <input type="text" name="telefon" class="form-control" value="{{ hasta[6] or '' }}">
                    </div>
                </div>

                <div class="mb-3">
                    <label class="form-label fw-semibold">E-Posta Adresi</label>
                    <input type="email" name="email" class="form-control" value="{{ hasta[7] or '' }}">
                </div>

                <div class="mb-4">
                    <label class="form-label fw-semibold">İkametgah Adresi</label>
                    <textarea name="adres" class="form-control" rows="2">{{ hasta[10] or '' }}</textarea>
                </div>

                <div class="d-flex gap-2">
                    <a href="/hasta/{{ hasta[0] }}" class="btn btn-outline-secondary w-50 rounded-pill">Vazgeç</a>
                    <button type="submit" class="btn btn-warning w-50 fw-bold rounded-pill text-dark shadow-sm">Değişiklikleri Kaydet</button>
                </div>
            </form>
        </div>
    </div>
</div>
"""

HEKIM_EKLE_CONTENT = """
<div class="row justify-content-center">
    <div class="col-md-8 col-lg-7">
        <div class="card card-custom p-4 p-md-5">
            <div class="text-center mb-4">
                <div class="rounded-circle bg-warning-subtle text-warning-emphasis p-3 d-inline-flex mb-2">
                    <i class="bi bi-person-badge-fill fs-3"></i>
                </div>
                <h4 class="fw-bold">Yeni Hekim Kayıt Paneli</h4>
                <p class="text-muted small">Poliklinik hekim kadrosuna yeni diş hekimi ekleme formu</p>
            </div>
            
            <form method="POST">
                <div class="row g-3 mb-3">
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">Adı</label>
                        <input type="text" name="ad" class="form-control" placeholder="Hekim Adı" required>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">Soyadı</label>
                        <input type="text" name="soyad" class="form-control" placeholder="Hekim Soyadı" required>
                    </div>
                </div>

                <div class="mb-3">
                    <label class="form-label fw-semibold">Uzmanlık Branşı</label>
                    <select name="brans_id" class="form-select" required>
                        <option value="">-- Branş Seçiniz --</option>
                        {% for b in branslar %}
                        <option value="{{ b[0] }}">{{ b[1] }}</option>
                        {% endfor %}
                    </select>
                </div>

                <div class="row g-3 mb-3">
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">Oda Numarası</label>
                        <input type="text" name="oda_no" class="form-control" placeholder="Örn: 204, B-12" required>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">Telefon Numarası</label>
                        <input type="text" name="telefon" class="form-control" placeholder="05XXXXXXXXX" required>
                    </div>
                </div>

                <div class="mb-4">
                    <label class="form-label fw-semibold">Kurumsal E-Posta Adresi</label>
                    <input type="email" name="email" class="form-control" placeholder="ad.soyad@dishastanesi.com" required>
                </div>

                <div class="d-flex gap-2">
                    <a href="/" class="btn btn-outline-secondary w-50 rounded-pill">İptal</a>
                    <button type="submit" class="btn btn-warning w-50 fw-bold rounded-pill text-dark shadow-sm">Hekimi Kadroya Kaydet</button>
                </div>
            </form>
        </div>
    </div>
</div>
"""

STOKLAR_CONTENT = """
<div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-4">
    <div>
        <h4 class="fw-bold mb-1"><i class="bi bi-box-seam-fill text-primary me-2"></i>Klinik Sarf Malzeme & Stok Envanteri</h4>
        <span class="text-muted small">Malzeme arama, stok takibi ve doğrudan elle miktar güncelleme paneli</span>
    </div>
    <div class="d-flex gap-2">
        <form class="d-flex" action="/stoklar" method="GET">
            <input class="form-control form-control-sm rounded-start-pill px-3" type="search" name="q" value="{{ arama or '' }}" placeholder="Malzeme veya Kategori Ara...">
            <button class="btn btn-primary btn-sm rounded-end-pill px-3" type="submit"><i class="bi bi-search"></i></button>
        </form>
        <button type="button" class="btn btn-primary btn-sm rounded-pill px-3" data-bs-toggle="modal" data-bs-target="#yeniMalzemeModal">
            <i class="bi bi-plus-lg me-1"></i> Yeni Malzeme Ekle
        </button>
        <a href="/" class="btn btn-outline-secondary btn-sm rounded-pill"><i class="bi bi-arrow-left me-1"></i> Ana Sayfa</a>
    </div>
</div>

<div class="card card-custom p-4">
    <div class="table-responsive">
        <table class="table table-hover align-middle">
            <thead class="table-light">
                <tr>
                    <th>#ID</th>
                    <th>Malzeme Adı</th>
                    <th>Kategori</th>
                    <th>Mevcut Miktar</th>
                    <th>Kritik Eşik</th>
                    <th>Durum</th>
                    <th>Miktar Güncelle (Elle Giriş)</th>
                    <th class="text-end">İşlem</th>
                </tr>
            </thead>
            <tbody>
                {% for s in stoklar %}
                <tr class="{% if s[3] <= s[5] %}table-danger-subtle{% endif %}">
                    <td><span class="badge bg-secondary-subtle text-body">{{ s[0] }}</span></td>
                    <td class="fw-bold">{{ s[1] }}</td>
                    <td><span class="badge bg-light text-secondary border">{{ s[2] }}</span></td>
                    <td class="fs-6 fw-bold text-primary">{{ s[3] }} <span class="small text-muted fw-normal">{{ s[4] }}</span></td>
                    <td>{{ s[5] }} {{ s[4] }}</td>
                    <td>
                        {% if s[3] <= 0 %}
                            <span class="badge bg-danger text-white"><i class="bi bi-x-circle me-1"></i>TÜKENDİ</span>
                        {% elif s[3] <= s[5] %}
                            <span class="badge bg-warning text-dark"><i class="bi bi-exclamation-triangle me-1"></i>KRİTİK STOK</span>
                        {% else %}
                            <span class="badge bg-success-subtle text-success"><i class="bi bi-check-circle me-1"></i>Yeterli</span>
                        {% endif %}
                    </td>
                    <td>
                        <form action="/stok-manuel-guncelle" method="POST" class="d-flex align-items-center gap-1" style="max-width: 170px;">
                            <input type="hidden" name="malzeme_id" value="{{ s[0] }}">
                            <input type="number" name="yeni_miktar" class="form-control form-control-sm text-center fw-bold" value="{{ s[3] }}" min="0" required>
                            <button type="submit" class="btn btn-outline-primary btn-sm rounded px-2" title="Miktarı Kaydet">
                                <i class="bi bi-check-lg"></i>
                            </button>
                        </form>
                    </td>
                    <td class="text-end">
                        <div class="btn-group btn-group-sm rounded-pill" role="group">
                            <a href="/stok-hareket/{{ s[0] }}/azalt" class="btn btn-outline-danger fw-bold" title="1 Düş">-1</a>
                            <a href="/stok-hareket/{{ s[0] }}/arttir" class="btn btn-outline-success fw-bold" title="1 Ekle">+1</a>
                            <a href="/stok-sil/{{ s[0] }}" class="btn btn-outline-secondary" onclick="return confirm('Bu malzemeyi silmek istediğinize emin misiniz?');">
                                <i class="bi bi-trash"></i>
                            </a>
                        </div>
                    </td>
                </tr>
                {% else %}
                <tr>
                    <td colspan="8" class="text-center text-muted py-4">Aranan kriterde stok kaydı bulunamadı.</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>

<div class="modal fade" id="yeniMalzemeModal" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content rounded-4 border-0 shadow">
            <div class="modal-header border-bottom">
                <h5 class="modal-title fw-bold"><i class="bi bi-box-seam text-primary me-2"></i>Yeni Sarf Malzeme Kaydı</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form action="/stok-ekle" method="POST">
                <div class="modal-body">
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Malzeme Adı</label>
                        <input type="text" name="malzeme_adi" class="form-control" placeholder="Örn: Ortodontik Tel, Rulo Pamuk, Kompozit" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Kategori</label>
                        <select name="kategori" class="form-select" required>
                            <option value="Sarf Malzeme">Sarf Malzeme (Pamuk, Maske, Eldiven vb.)</option>
                            <option value="Ortodonti">Ortodonti (Tel, Braket, Elastik)</option>
                            <option value="Endodonti">Endodonti (Kanal Eğesi, Solüsyon)</option>
                            <option value="Restoratif">Restoratif (Kompozit, Bonding)</option>
                            <option value="Cerrahi">Cerrahi & Anestezi (Ampul, İğne Ucu)</option>
                            <option value="Protez">Protez Malzemeleri</option>
                        </select>
                    </div>
                    <div class="row g-2 mb-3">
                        <div class="col-md-6">
                            <label class="form-label fw-semibold">Başlangıç Miktarı</label>
                            <input type="number" name="miktar" class="form-control" value="10" min="0" required>
                        </div>
                        <div class="col-md-6">
                            <label class="form-label fw-semibold">Ölçü Birimi</label>
                            <select name="birim" class="form-select" required>
                                <option value="Adet">Adet</option>
                                <option value="Paket">Paket</option>
                                <option value="Kutu">Kutu</option>
                                <option value="Rulo">Rulo</option>
                                <option value="Tüp">Tüp</option>
                                <option value="Set">Set</option>
                                <option value="Ampul">Ampul</option>
                            </select>
                        </div>
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Kritik Stok Uyarısı Eşiği</label>
                        <input type="number" name="kritik_seviye" class="form-control" value="10" min="1" required>
                    </div>
                </div>
                <div class="modal-footer border-top">
                    <button type="button" class="btn btn-secondary rounded-pill" data-bs-dismiss="modal">İptal</button>
                    <button type="submit" class="btn btn-primary fw-bold rounded-pill px-4">Stoğa Ekle</button>
                </div>
            </form>
        </div>
    </div>
</div>
"""

HASTA_DETAY_CONTENT = """
<div class="d-flex justify-content-between align-items-center mb-4">
    <div>
        <h4 class="fw-bold mb-1"><i class="bi bi-person-lines-fill text-primary me-2"></i>{{ hasta[2] }} {{ hasta[3] }}</h4>
        <div class="d-flex align-items-center gap-2 mt-1">
            <span class="badge bg-primary-subtle text-primary fw-bold px-3 py-1">Hasta No: {{ hasta[9] or ('HST-' ~ hasta[0]) }}</span>
            <a href="/hastalar" class="btn btn-sm btn-outline-secondary rounded-pill px-2 py-0" style="font-size: 12px;"><i class="bi bi-arrow-left me-1"></i>Geri</a>
        </div>
    </div>
    <div class="d-flex flex-wrap gap-2">
        <button type="button" class="btn btn-sm text-white rounded-pill px-3 fw-bold shadow-sm" style="background-color: #7c3aed;" data-bs-toggle="modal" data-bs-target="#planModal">
            <i class="bi bi-clipboard2-pulse me-1"></i> + Tedavi Planı / Teklif
        </button>
        <button type="button" class="btn btn-primary btn-sm rounded-pill px-3 fw-bold shadow-sm" data-bs-toggle="modal" data-bs-target="#tedaviModal">
            <i class="bi bi-capsule me-1"></i> + Tedavi Uygula
        </button>
        <button type="button" class="btn btn-success btn-sm rounded-pill px-3 fw-bold shadow-sm" data-bs-toggle="modal" data-bs-target="#odemeModal">
            <i class="bi bi-cash-stack me-1"></i> + Ödeme Al
        </button>
    </div>
</div>

{% if hasta[8] %}
<div class="alert alert-danger d-flex align-items-center rounded-4 shadow-sm mb-4" role="alert">
    <i class="bi bi-exclamation-triangle-fill fs-3 me-3 text-danger"></i>
    <div>
        <h6 class="alert-heading fw-bold mb-1">DİKKAT: HASTA TIBBİ UYARISI / ALERJİ / DİYABET NOTU</h6>
        <p class="mb-0 small fw-semibold">{{ hasta[8] }}</p>
    </div>
</div>
{% endif %}

<div class="row g-3 mb-4">
    <div class="col-md-4">
        <div class="card card-custom p-3 bg-light border-0 shadow-sm">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <span class="text-muted small fw-bold text-uppercase d-block">Toplam Hizmet Tutarı</span>
                    <h4 class="fw-bold text-dark mb-0">{{ "%.2f"|format(toplam_tedavi_tutari) }} ₺</h4>
                </div>
                <div class="p-3 bg-primary-subtle text-primary rounded-circle">
                    <i class="bi bi-receipt fs-4"></i>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card card-custom p-3 bg-light border-0 shadow-sm">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <span class="text-muted small fw-bold text-uppercase d-block">Tahsil Edilen Ödemeler</span>
                    <h4 class="fw-bold text-success mb-0">{{ "%.2f"|format(toplam_odenen) }} ₺</h4>
                </div>
                <div class="p-3 bg-success-subtle text-success rounded-circle">
                    <i class="bi bi-wallet2 fs-4"></i>
                </div>
            </div>
        </div>
    </div>
    <div class="col-md-4">
        <div class="card card-custom p-3 {% if kalan_borc > 0 %}bg-danger-subtle border-danger{% else %}bg-success-subtle border-success{% endif %} border shadow-sm">
            <div class="d-flex justify-content-between align-items-center">
                <div>
                    <span class="text-muted small fw-bold text-uppercase d-block">Kalan Borç / Bakiye</span>
                    <h4 class="fw-bold {% if kalan_borc > 0 %}text-danger{% else %}text-success{% endif %} mb-0">
                        {{ "%.2f"|format(kalan_borc) }} ₺
                    </h4>
                </div>
                <div>
                    {% if kalan_borc > 0 %}
                        <span class="badge bg-danger text-white"><i class="bi bi-exclamation-circle me-1"></i>Borç Var</span>
                    {% else %}
                        <span class="badge bg-success text-white"><i class="bi bi-check-circle me-1"></i>Borcu Yok</span>
                    {% endif %}
                </div>
            </div>
        </div>
    </div>
</div>

<div class="card card-custom p-4 mb-4">
    <div class="d-flex justify-content-between align-items-center mb-3">
        <div>
            <h5 class="fw-bold mb-0" style="color: #7c3aed;"><i class="bi bi-clipboard2-pulse me-2"></i>Aktif Tedavi Planları & Seans Takibi</h5>
            <small class="text-muted">Aşamalı tedavi planları, seans ilerleme durumları ve fiyat teklifleri</small>
        </div>
        <button type="button" class="btn btn-outline-primary btn-sm rounded-pill" data-bs-toggle="modal" data-bs-target="#planModal">
            <i class="bi bi-plus-lg me-1"></i> Yeni Plan Oluştur
        </button>
    </div>
    
    <div class="row g-3">
        {% for p in planlar %}
        <div class="col-lg-6">
            <div class="card card-custom border p-3 h-100 bg-body-tertiary">
                <div class="d-flex justify-content-between align-items-start mb-2">
                    <div>
                        <span class="badge bg-primary-subtle text-primary border border-primary-subtle fw-bold me-1">PLN-{{ "%04d"|format(p[0]) }}</span>
                        <strong class="fs-6">{{ p[3] }}</strong>
                    </div>
                    <div>
                        {% if p[7] == 'Tamamlandı' %}
                            <span class="badge bg-success text-white">Tamamlandı</span>
                        {% else %}
                            <span class="badge bg-warning text-dark">{{ p[7] }}</span>
                        {% endif %}
                    </div>
                </div>
                
                <p class="small text-muted mb-2">{{ p[4] or 'Açıklama belirtilmedi.' }}</p>

                <div class="mb-3">
                    <div class="d-flex justify-content-between small fw-semibold mb-1">
                        <span>Seans İlerlemesi:</span>
                        <span>{{ p[6] }} / {{ p[5] }} Seans (%{{ ((p[6] / p[5]) * 100)|round|int if p[5] > 0 else 0 }})</span>
                    </div>
                    <div class="progress" style="height: 10px;">
                        <div class="progress-bar bg-primary" role="progressbar" style="width: {{ ((p[6] / p[5]) * 100)|round|int if p[5] > 0 else 0 }}%;"></div>
                    </div>
                </div>

                <div class="d-flex justify-content-between align-items-center pt-2 border-top small">
                    <div>
                        <span class="text-muted">Hekim: <strong>Dt. {{ p[10] }}</strong></span><br>
                        <span class="text-muted">Tahmini Tutar: <strong class="text-success">{{ "%.2f"|format(p[2]) }} ₺</strong></span>
                    </div>
                    <div class="d-flex gap-1">
                        {% if p[6] < p[5] %}
                        <a href="/plan-seans-arttir/{{ p[0] }}?hasta_id={{ hasta[0] }}" class="btn btn-sm btn-success rounded-pill px-2" title="1 Seans İlerlet">
                            <i class="bi bi-check-lg me-1"></i>+1 Seans
                        </a>
                        {% endif %}
                        <a href="/plan-teklif-yazdir/{{ p[0] }}" target="_blank" class="btn btn-sm btn-outline-primary rounded-pill px-2" title="Fiyat Teklifi & Plan Çıktısı">
                            <i class="bi bi-printer"></i> Teklif
                        </a>
                        <a href="/plan-sil/{{ p[0] }}?hasta_id={{ hasta[0] }}" class="btn btn-sm btn-outline-danger rounded-pill px-2" onclick="return confirm('Bu tedavi planını silmek istiyor musunuz?');">
                            <i class="bi bi-trash"></i>
                        </a>
                    </div>
                </div>
            </div>
        </div>
        {% else %}
        <div class="col-12 text-center text-muted py-3">
            Hastaya tanımlanmış aktif bir tedavi planı bulunmuyor. Yeni bir plan açıp hastaya teklif sunabilirsiniz.
        </div>
        {% endfor %}
    </div>
</div>

<div class="card card-custom p-4 mb-4">
    <div class="row g-3">
        <div class="col-md-3"><span class="text-muted small d-block">Hasta Numarası</span><span class="fw-bold text-primary">{{ hasta[9] or ('HST-' ~ hasta[0]) }}</span></div>
        <div class="col-md-3"><span class="text-muted small d-block">T.C. Kimlik No</span><span class="fw-semibold">{{ hasta[1] or '-' }}</span></div>
        <div class="col-md-2"><span class="text-muted small d-block">Doğum Tarihi</span><span class="fw-semibold">{{ hasta[4] or '-' }}</span></div>
        <div class="col-md-2"><span class="text-muted small d-block">Cinsiyet</span><span class="fw-semibold">{{ hasta[5] }}</span></div>
        <div class="col-md-2"><span class="text-muted small d-block">Telefon</span><span class="fw-semibold">{{ hasta[6] or '-' }}</span></div>
    </div>
</div>

<div class="card card-custom p-4 mb-4">
    <div class="d-flex justify-content-between align-items-center mb-3">
        <h5 class="fw-bold mb-0"><i class="bi bi-capsule-pill text-primary me-2"></i>Uygulanan Tedaviler & İşlemler</h5>
        <button type="button" class="btn btn-outline-primary btn-sm rounded-pill" data-bs-toggle="modal" data-bs-target="#tedaviModal">
            <i class="bi bi-plus-lg me-1"></i> Yeni Tedavi
        </button>
    </div>
    <div class="table-responsive">
        <table class="table table-hover align-middle">
            <thead class="table-light">
                <tr><th>Tarih</th><th>Tanı / Teşhis</th><th>Diş No</th><th>İşlem Adı</th><th class="text-end">Tutar</th><th class="text-end">İşlem</th></tr>
            </thead>
            <tbody>
                {% for t in tedaviler %}
                <tr>
                    <td>{{ t[2] }}</td>
                    <td class="fw-semibold">{{ t[1] }}</td>
                    <td><span class="badge bg-dark-subtle text-body">{{ t[3] or '-' }}</span></td>
                    <td>{{ t[4] }}</td>
                    <td class="text-end fw-bold text-primary">{{ t[5] }} ₺</td>
                    <td class="text-end">
                        <a href="/tedavi-sil/{{ t[0] }}?hasta_id={{ hasta[0] }}" class="btn btn-outline-danger btn-sm rounded-pill px-2" onclick="return confirm('Bu tedavi kaydını silmek istediğinize emin misiniz?');">
                            <i class="bi bi-trash"></i>
                        </a>
                    </td>
                </tr>
                {% else %}
                <tr><td colspan="6" class="text-muted text-center py-3">Kayıtlı tedavi bulunmuyor.</td></tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>

<div class="card card-custom p-4 mb-4">
    <div class="d-flex justify-content-between align-items-center mb-3">
        <h5 class="fw-bold mb-0"><i class="bi bi-receipt-cutoff text-success me-2"></i>Finans & Ödeme Kayıtları</h5>
        <button type="button" class="btn btn-outline-success btn-sm rounded-pill" data-bs-toggle="modal" data-bs-target="#odemeModal">
            <i class="bi bi-plus-lg me-1"></i> Yeni Ödeme Al
        </button>
    </div>
    <table class="table table-hover align-middle">
        <thead><tr><th>Tarih</th><th>Yöntem</th><th>Durum</th><th>Açıklama</th><th class="text-end">Tutar</th><th class="text-end">Makbuz</th></tr></thead>
        <tbody>
            {% for o in odemeler %}
            <tr>
                <td>{{ o[1] }}</td>
                <td>{{ o[2] }}</td>
                <td>
                    {% if o[3] == 'Ödendi' %}<span class="badge bg-success-subtle text-success">Ödendi</span>
                    {% elif o[3] == 'Bekliyor' %}<span class="badge bg-warning-subtle text-warning-emphasis">Bekliyor</span>
                    {% else %}<span class="badge bg-danger-subtle text-danger">{{ o[3] }}</span>{% endif %}
                </td>
                <td class="small text-muted">{{ o[4] or '-' }}</td>
                <td class="text-end fw-bold text-success">{{ o[0] }} ₺</td>
                <td class="text-end">
                    <a href="/odeme-makbuz/{{ o[5] }}" target="_blank" class="btn btn-outline-primary btn-sm rounded-pill px-2 py-1">
                        <i class="bi bi-printer me-1"></i> Yazdır
                    </a>
                </td>
            </tr>
            {% else %}
            <tr><td colspan="6" class="text-muted text-center py-3">Ödeme kaydı bulunmuyor.</td></tr>
            {% endfor %}
        </tbody>
    </table>
</div>

<div class="card card-custom p-4 bg-body-tertiary border">
    <div class="d-flex flex-wrap justify-content-between align-items-center gap-3">
        <div>
            <h6 class="fw-bold mb-1"><i class="bi bi-gear-wide-connected text-secondary me-2"></i>Hasta Dosyası Yönetim Seçenekleri</h6>
            <small class="text-muted">Bu hastaya ait demografik bilgileri güncelleyebilir veya arşiv geçmişini inceleyebilirsiniz.</small>
        </div>
        <div class="d-flex gap-2">
            <a href="/hasta/{{ hasta[0] }}/duzenle" class="btn btn-warning btn-sm rounded-pill px-3 fw-bold text-dark shadow-sm">
                <i class="bi bi-pencil-square me-1"></i> Bilgileri Düzenle
            </a>
            <a href="/hasta/{{ hasta[0] }}/gecmis" class="btn btn-outline-info btn-sm rounded-pill px-3 fw-bold">
                <i class="bi bi-clock-history me-1"></i> Tüm Tıbbi Geçmiş
            </a>
            <a href="/hasta-sil/{{ hasta[0] }}" class="btn btn-outline-danger btn-sm rounded-pill px-3" onclick="return confirm('Bu hastayı ve tüm tıbbi geçmişini kalıcı olarak silmek istiyor musunuz?');">
                <i class="bi bi-trash me-1"></i> Hastayı Sil
            </a>
        </div>
    </div>
</div>

<!-- MODALLAR -->
<div class="modal fade" id="planModal" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content rounded-4 border-0 shadow">
            <div class="modal-header border-bottom">
                <h5 class="modal-title fw-bold" style="color: #7c3aed;"><i class="bi bi-clipboard2-pulse me-2"></i>Yeni Tedavi Planı & Fiyat Teklifi</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form action="/hasta/{{ hasta[0] }}/plan-ekle" method="POST">
                <div class="modal-body">
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Tedavi Planı Adı</label>
                        <input type="text" name="plan_adi" class="form-control" placeholder="Örn: 4 Seanslık Ortodonti Tedavisi, 6 Üye Zirkonyum Kaplama" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Sorumlu Diş Hekimi</label>
                        <select name="hekim_id" class="form-select" required>
                            {% for d in tum_hekimler %}
                            <option value="{{ d[0] }}">Dt. {{ d[1] }} {{ d[2] }}</option>
                            {% endfor %}
                        </select>
                    </div>
                    <div class="row g-2 mb-3">
                        <div class="col-md-6">
                            <label class="form-label fw-semibold">Öngörülen Seans Sayısı</label>
                            <input type="number" name="toplam_seans" class="form-control" value="3" min="1" required>
                        </div>
                        <div class="col-md-6">
                            <label class="form-label fw-semibold">Tahmini Paket Tutarı (₺)</label>
                            <input type="number" step="0.01" name="tahmini_tutar" class="form-control" placeholder="Örn: 8500.00" required>
                        </div>
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Tedavi Planı & Seans Notları</label>
                        <textarea name="aciklama" class="form-control" rows="3" placeholder="1. Seans: Çürük temizleme ve dolgu. 2. Seans: Ölçü alma..."></textarea>
                    </div>
                </div>
                <div class="modal-footer border-top">
                    <button type="button" class="btn btn-secondary rounded-pill" data-bs-dismiss="modal">İptal</button>
                    <button type="submit" class="btn btn-primary fw-bold rounded-pill px-4" style="background-color: #7c3aed; border-color: #7c3aed;">Planı Kaydet & Teklif Oluştur</button>
                </div>
            </form>
        </div>
    </div>
</div>

<div class="modal fade" id="tedaviModal" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content rounded-4 border-0 shadow">
            <div class="modal-header border-bottom">
                <h5 class="modal-title fw-bold"><i class="bi bi-capsule-pill text-primary me-2"></i>Tedavi & İşlem Ekle</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form action="/hasta/{{ hasta[0] }}/tedavi-ekle" method="POST">
                <div class="modal-body">
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Tanı / Teşhis</label>
                        <input type="text" name="tani" class="form-control" placeholder="Örn: Pulpa Nekrozu" required>
                    </div>
                    <div class="row g-2 mb-3">
                        <div class="col-md-7">
                            <label class="form-label fw-semibold">Yapılan İşlem</label>
                            <select name="islem_id" class="form-select" required>
                                {% for dis in dis_islemleri %}
                                <option value="{{ dis[0] }}">{{ dis[1] }}</option>
                                {% endfor %}
                            </select>
                        </div>
                        <div class="col-md-5">
                            <label class="form-label fw-semibold">Diş Numarası</label>
                            <input type="text" name="dis_no" class="form-control fw-bold text-primary" placeholder="Örn: 24, 46">
                        </div>
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-semibold">İşlem Ücreti (₺)</label>
                        <input type="number" step="0.01" name="fiyat" class="form-control" placeholder="Örn: 1500.00" required>
                    </div>
                </div>
                <div class="modal-footer border-top">
                    <button type="button" class="btn btn-secondary rounded-pill" data-bs-dismiss="modal">Vazgeç</button>
                    <button type="submit" class="btn btn-primary fw-bold rounded-pill px-4">Kaydet</button>
                </div>
            </form>
        </div>
    </div>
</div>

<div class="modal fade" id="odemeModal" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content rounded-4 border-0 shadow">
            <div class="modal-header border-bottom">
                <h5 class="modal-title fw-bold"><i class="bi bi-cash-stack text-success me-2"></i>Hasta Ödeme Tahsilatı</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form action="/hasta/{{ hasta[0] }}/odeme-ekle" method="POST">
                <div class="modal-body">
                    <div class="alert alert-secondary py-2 small mb-3 d-flex justify-content-between align-items-center">
                        <span>Mevcut Kalan Borç:</span>
                        <strong class="{% if kalan_borc > 0 %}text-danger fs-6{% else %}text-success{% endif %}">
                            {{ "%.2f"|format(kalan_borc) }} ₺
                        </strong>
                    </div>

                    <div class="mb-3">
                        <label class="form-label fw-semibold">Tahsil Edilen Tutar (₺)</label>
                        <input type="number" step="0.01" name="tutar" class="form-control form-control-lg fw-bold text-success" value="{{ "%.2f"|format(kalan_borc) if kalan_borc > 0 else '' }}" placeholder="0.00" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Ödeme Yöntemi</label>
                        <select name="odeme_yontemi" class="form-select" required>
                            <option value="Kredi Kartı">Kredi Kartı</option>
                            <option value="Nakit">Nakit</option>
                            <option value="Banka Kartı (POS)">Banka Kartı (POS)</option>
                            <option value="Havale">Havale / EFT</option>
                        </select>
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Ödeme Durumu</label>
                        <select name="durum" class="form-select" required>
                            <option value="Ödendi">Ödendi (Tahsil Edildi)</option>
                            <option value="Bekliyor">Bekliyor (Onay Aşamasında)</option>
                        </select>
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Açıklama / Makbuz Notu</label>
                        <textarea name="aciklama" class="form-control" rows="2" placeholder="Örn: Tedavi ara ödemesi"></textarea>
                    </div>
                </div>
                <div class="modal-footer border-top">
                    <button type="button" class="btn btn-secondary rounded-pill" data-bs-dismiss="modal">İptal</button>
                    <button type="submit" class="btn btn-success fw-bold rounded-pill px-4">Ödemeyi Kaydet & Düş</button>
                </div>
            </form>
        </div>
    </div>
</div>
"""

HASTA_GECMIS_CONTENT = """
<div class="d-flex justify-content-between align-items-center mb-4">
    <div>
        <h4 class="fw-bold mb-1"><i class="bi bi-clock-history text-info me-2"></i>Hasta Tıbbi Arşivi: {{ hasta[2] }} {{ hasta[3] }}</h4>
        <span class="text-muted small">Hasta No: <strong>{{ hasta[9] or ('HST-' ~ hasta[0]) }}</strong> | T.C: {{ hasta[1] or '-' }} | Telefon: {{ hasta[6] or '-' }}</span>
    </div>
    <div>
        <a href="/hasta/{{ hasta[0] }}" class="btn btn-primary btn-sm rounded-pill px-3 me-2"><i class="bi bi-file-earmark-medical me-1"></i> Dosyaya Dön</a>
        <a href="/hastalar" class="btn btn-outline-secondary btn-sm rounded-pill"><i class="bi bi-arrow-left me-1"></i> Hasta Listesi</a>
    </div>
</div>

{% if hasta[8] %}
<div class="alert alert-danger rounded-4 py-2 mb-4">
    <i class="bi bi-exclamation-triangle-fill me-1"></i> <strong>Tıbbi Uyarı Notu:</strong> {{ hasta[8] }}
</div>
{% endif %}

<div class="row g-4">
    <div class="col-12">
        <div class="card card-custom p-4">
            <h5 class="fw-bold mb-3"><i class="bi bi-calendar-event text-primary me-2"></i>Randevu Çizelgesi Geçmişi</h5>
            <div class="table-responsive">
                <table class="table table-hover align-middle">
                    <thead class="table-light"><tr><th>Tarih</th><th>Saat</th><th>Hekim</th><th>Durum</th><th>Şikayet / Not</th></tr></thead>
                    <tbody>
                        {% for r in randevular %}
                        <tr>
                            <td class="fw-semibold">{{ r[0] }}</td>
                            <td>{{ r[1] }}</td>
                            <td>Dt. {{ r[2] }}</td>
                            <td><span class="badge bg-secondary-subtle text-body">{{ r[3] }}</span></td>
                            <td class="small text-muted">{{ r[4] or '-' }}</td>
                        </tr>
                        {% else %}
                        <tr><td colspan="5" class="text-center text-muted py-3">Randevu kaydı bulunamadı.</td></tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <div class="col-12">
        <div class="card card-custom p-4 h-100">
            <h5 class="fw-bold mb-3"><i class="bi bi-shield-shaded text-success me-2"></i>Uygulanan Tedavi Kayıtları</h5>
            <ul class="list-group list-group-flush">
                {% for t in tedaviler %}
                <li class="list-group-item px-0">
                    <div class="d-flex justify-content-between align-items-center">
                        <strong>{{ t[1] }}</strong>
                        <span class="badge bg-success-subtle text-success">{{ t[4] }} ₺</span>
                    </div>
                    <div class="small text-muted">İşlem: {{ t[3] }} (Diş: {{ t[2] or 'Genel' }}) - Tarih: {{ t[0] }}</div>
                </li>
                {% else %}
                <li class="list-group-item px-0 text-muted">Geçmiş tedavi bulunmuyor.</li>
                {% endfor %}
            </ul>
        </div>
    </div>
</div>
"""

RANDEVULAR_CONTENT = """
<div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-4">
    <div>
        <h4 class="fw-bold mb-1"><i class="bi bi-calendar-range-fill text-primary me-2"></i>Dt. {{ hekim[1] }} {{ hekim[2] }}</h4>
        <span class="text-muted small">Klinik Randevu Çizelgesi & Günlük Hatırlatma Özeti</span>
    </div>
    <div class="d-flex gap-2">
        <a href="/hekim/{{ hekim[0] }}/bugun-ozet-gonder" class="btn btn-warning text-dark btn-sm rounded-pill px-3 fw-bold shadow-sm">
            <i class="bi bi-envelope-check-fill me-1"></i> Bugünkü Tüm Randevuları Hatırlat (Günlük Özet)
        </a>
        <a href="/" class="btn btn-outline-secondary btn-sm rounded-pill"><i class="bi bi-arrow-left me-1"></i> Ana Sayfa</a>
    </div>
</div>

<div class="card card-custom p-4">
    <div class="table-responsive">
        <table class="table table-hover align-middle">
            <thead class="table-light">
                <tr>
                    <th>Tarih</th>
                    <th>Saat</th>
                    <th>Hasta Bilgisi</th>
                    <th>İletişim</th>
                    <th>Durum</th>
                    <th>Not / Şikayet</th>
                    <th class="text-end">İşlemler</th>
                </tr>
            </thead>
            <tbody>
                {% for r in randevular %}
                <tr>
                    <td class="fw-bold">{{ r[4] }}</td>
                    <td><span class="badge bg-secondary-subtle text-body">{{ r[5] }}</span></td>
                    <td class="fw-semibold">{{ r[2] }}</td>
                    <td class="small text-muted">{{ r[3] or '-' }}</td>
                    <td>
                        {% if r[6] == 'Tamamlandı' %}<span class="badge bg-success-subtle text-success">Tamamlandı</span>
                        {% elif r[6] == 'Bekliyor' %}<span class="badge bg-warning-subtle text-warning-emphasis">Bekliyor</span>
                        {% else %}<span class="badge bg-danger-subtle text-danger">{{ r[6] }}</span>{% endif %}
                    </td>
                    <td class="small text-muted">{{ r[7] or '-' }}</td>
                    <td class="text-end">
                        <button type="button" class="btn btn-outline-warning btn-sm rounded-pill px-2 me-1" 
                                data-bs-toggle="modal" 
                                data-bs-target="#mailModal"
                                data-hekim-id="{{ hekim[0] }}"
                                data-hekim-ad="Dt. {{ hekim[1] }} {{ hekim[2] }}"
                                data-hekim-mail="{{ hekim[3] }}"
                                data-hasta-ad="{{ r[2] }}"
                                data-randevu-tarih="{{ r[4] }}"
                                data-randevu-saat="{{ r[5] }}">
                            <i class="bi bi-envelope-at me-1"></i> Hatırlat
                        </button>
                        <a href="/hasta/{{ r[1] }}" class="btn btn-primary btn-sm rounded-pill px-3 me-1">
                            <i class="bi bi-file-medical me-1"></i> Dosya
                        </a>
                        <a href="/randevu-sil/{{ r[0] }}?hekim_id={{ hekim[0] }}" class="btn btn-outline-danger btn-sm rounded-pill px-2" onclick="return confirm('Randevuyu silmek istediğinize emin misiniz?');">
                            <i class="bi bi-trash"></i>
                        </a>
                    </td>
                </tr>
                {% else %}
                <tr>
                    <td colspan="7" class="text-center text-muted py-5">Bu hekime tanımlanmış randevu bulunamadı.</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>

<div class="modal fade" id="mailModal" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog">
        <div class="modal-content rounded-4 border-0 shadow">
            <div class="modal-header border-bottom">
                <h5 class="modal-title fw-bold"><i class="bi bi-chat-dots-fill text-warning me-2"></i>Randevu Hatırlatma Bildirimi</h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <form action="/hekim-mail-hatirlat" method="POST">
                <input type="hidden" name="hekim_id" id="modalHekimId">
                <input type="hidden" name="hekim_mail" id="modalHekimMail">
                <div class="modal-body">
                    <div class="alert alert-info py-2 small mb-3">
                        Hekim: <strong id="modalHekimAd"></strong> (<span id="modalHekimMailDisplay"></span>)
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Bildirim Saati</label>
                        <input type="time" name="hatirlatma_saati" class="form-control" value="08:30" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label fw-semibold">Bildirim Metni</label>
                        <textarea class="form-control bg-body-secondary small" rows="3" readonly id="modalKlasikMesaj"></textarea>
                    </div>
                </div>
                <div class="modal-footer border-top">
                    <button type="button" class="btn btn-secondary rounded-pill" data-bs-dismiss="modal">İptal</button>
                    <button type="submit" class="btn btn-warning fw-bold rounded-pill px-4 text-dark">Gönder</button>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
    const mailModal = document.getElementById('mailModal');
    if (mailModal) {
        mailModal.addEventListener('show.bs.modal', function (event) {
            const button = event.relatedTarget;
            const hekimId = button.getAttribute('data-hekim-id');
            const hekimAd = button.getAttribute('data-hekim-ad');
            const hekimMail = button.getAttribute('data-hekim-mail');
            const hastaAd = button.getAttribute('data-hasta-ad');
            const tarih = button.getAttribute('data-randevu-tarih');
            const saat = button.getAttribute('data-randevu-saat');

            document.getElementById('modalHekimId').value = hekimId;
            document.getElementById('modalHekimMail').value = hekimMail;
            document.getElementById('modalHekimAd').innerText = hekimAd;
            document.getElementById('modalHekimMailDisplay').innerText = hekimMail;
            document.getElementById('modalKlasikMesaj').value = `Sayın ${hekimAd}, ${tarih} tarihinde saat ${saat} için hastanız ${hastaAd} ile olan randevunuz planlanmıştır.`;
        });
    }
</script>
"""

YENI_RANDEVU_CONTENT = """
<div class="row justify-content-center">
    <div class="col-md-8 col-lg-6">
        <div class="card card-custom p-4 p-md-5">
            <div class="text-center mb-4">
                <div class="rounded-circle bg-primary-subtle text-primary p-3 d-inline-flex mb-2">
                    <i class="bi bi-calendar-plus-fill fs-3"></i>
                </div>
                <h4 class="fw-bold">Yeni Randevu Kaydı</h4>
                <p class="text-muted small">Lütfen hasta ve poliklinik hekimini seçin</p>
            </div>
            
            <form method="POST">
                <div class="mb-3">
                    <label class="form-label fw-semibold">Hasta Seçimi</label>
                    <select name="hasta_id" class="form-select" required>
                        <option value="">-- Listeden Hasta Seçiniz --</option>
                        {% for h in hastalar %}
                        <option value="{{ h[0] }}">[{{ h[5] or ('HST-' ~ h[0]) }}] {{ h[1] }} {{ h[2] }} {% if h[3] %}(TC: {{ h[3] }}){% endif %} {% if h[4] %}- [UYARI: {{ h[4] }}]{% endif %}</option>
                        {% endfor %}
                    </select>
                </div>
                <div class="mb-3">
                    <label class="form-label fw-semibold">Diş Hekimi</label>
                    <select name="hekim_id" class="form-select" required>
                        <option value="">-- Görevli Hekim Seçiniz --</option>
                        {% for d in hekimler %}
                        <option value="{{ d[0] }}">Dt. {{ d[1] }} {{ d[2] }}</option>
                        {% endfor %}
                    </select>
                </div>
                <div class="row g-2 mb-3">
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">Randevu Tarihi</label>
                        <input type="date" name="tarih" class="form-control" required>
                    </div>
                    <div class="col-md-6">
                        <label class="form-label fw-semibold">Randevu Saati</label>
                        <input type="time" name="saat" class="form-control" required>
                    </div>
                </div>
                <div class="mb-4">
                    <label class="form-label fw-semibold">Şikayet & Ön Not</label>
                    <textarea name="notlar" class="form-control" rows="3" placeholder="Örn: Ağrı, kontrol, dolgu..."></textarea>
                </div>
                <div class="d-flex gap-2">
                    <a href="/" class="btn btn-outline-secondary w-50 rounded-pill">İptal</a>
                    <button type="submit" class="btn btn-primary w-50 fw-bold rounded-pill shadow-sm">Kaydı Tamamla</button>
                </div>
            </form>
        </div>
    </div>
</div>
"""

# ==========================================
# FLASK ROUTES
# ==========================================

@app.route('/')
def index():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT h.hekim_id, h.ad, h.soyad, b.brans_adi, h.oda_no, h.telefon
        FROM DisHekimleri h
        INNER JOIN Branslar b ON h.brans_id = b.brans_id
    """)
    hekimler = cursor.fetchall()
    
    cursor.execute("SELECT COUNT(*) FROM Hastalar")
    toplam_hasta = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM Randevular WHERE randevu_tarihi = CAST(GETDATE() AS DATE)")
    bugunku_randevu = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM Stoklar WHERE miktar <= kritik_seviye")
    kritik_stok_sayisi = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM Tedaviler")
    toplam_tedavi_sayisi = cursor.fetchone()[0]

    conn.close()

    content = render_template_string(INDEX_CONTENT, 
                                   hekimler=hekimler, 
                                   toplam_hasta=toplam_hasta, 
                                   bugunku_randevu=bugunku_randevu,
                                   kritik_stok_sayisi=kritik_stok_sayisi,
                                   toplam_tedavi_sayisi=toplam_tedavi_sayisi)
    return render_template_string(BASE_HTML, content=content)

@app.route('/kasa-raporlari')
def kasa_raporlari():
    filtre = request.args.get('filtre', 'aylik')
    baslangic = request.args.get('baslangic', '')
    bitis = request.args.get('bitis', '')

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT ISNULL(SUM(tutar), 0), COUNT(*) 
        FROM Odemeler 
        WHERE durum = N'Ödendi' AND CAST(odeme_tarihi AS DATE) = CAST(GETDATE() AS DATE)
    """)
    row_gun = cursor.fetchone()
    gunluk_tutar = float(row_gun[0] or 0.0)
    gunluk_adet = row_gun[1]

    cursor.execute("""
        SELECT ISNULL(SUM(tutar), 0), COUNT(*) 
        FROM Odemeler 
        WHERE durum = N'Ödendi' AND DATEDIFF(day, odeme_tarihi, GETDATE()) <= 7
    """)
    row_hafta = cursor.fetchone()
    haftalik_tutar = float(row_hafta[0] or 0.0)
    haftalik_adet = row_hafta[1]

    cursor.execute("""
        SELECT ISNULL(SUM(tutar), 0), COUNT(*) 
        FROM Odemeler 
        WHERE durum = N'Ödendi' AND DATEDIFF(day, odeme_tarihi, GETDATE()) <= 30
    """)
    row_ay = cursor.fetchone()
    aylik_tutar = float(row_ay[0] or 0.0)
    aylik_adet = row_ay[1]

    cursor.execute("""
        SELECT odeme_yontemi, SUM(tutar) 
        FROM Odemeler 
        WHERE durum = N'Ödendi' 
        GROUP BY odeme_yontemi
    """)
    grafik_satirlari = cursor.fetchall()
    grafik_etiketleri = [g[0] for g in grafik_satirlari] if grafik_satirlari else ["Kayıt Yok"]
    grafik_verileri = [float(g[1]) for g in grafik_satirlari] if grafik_satirlari else [0]

    sql_base = """
        SELECT o.odeme_id, CONVERT(VARCHAR(16), o.odeme_tarihi, 120), o.odeme_yontemi, o.tutar, o.aciklama,
               h.hasta_id, h.hasta_no, h.ad, h.soyad
        FROM Odemeler o
        INNER JOIN Hastalar h ON o.hasta_id = h.hasta_id
        WHERE o.durum = N'Ödendi'
    """
    params = []

    if filtre == 'gunluk':
        sql_base += " AND CAST(o.odeme_tarihi AS DATE) = CAST(GETDATE() AS DATE) ORDER BY o.odeme_tarihi DESC"
        filtre_aciklama = "Bugünün Tahsilatları"
        filtrelenen_toplam = gunluk_tutar
    elif filtre == 'haftalik':
        sql_base += " AND DATEDIFF(day, o.odeme_tarihi, GETDATE()) <= 7 ORDER BY o.odeme_tarihi DESC"
        filtre_aciklama = "Son 7 Gün (Haftalık Ciro)"
        filtrelenen_toplam = haftalik_tutar
    elif filtre == 'ozel' and baslangic and bitis:
        sql_base += " AND CAST(o.odeme_tarihi AS DATE) >= ? AND CAST(o.odeme_tarihi AS DATE) <= ? ORDER BY o.odeme_tarihi DESC"
        params = [baslangic, bitis]
        filtre_aciklama = f"{baslangic} ile {bitis} Arası"
        cursor.execute("SELECT ISNULL(SUM(tutar), 0) FROM Odemeler WHERE durum = N'Ödendi' AND CAST(odeme_tarihi AS DATE) >= ? AND CAST(odeme_tarihi AS DATE) <= ?", (baslangic, bitis))
        filtrelenen_toplam = float(cursor.fetchone()[0] or 0.0)
    else:
        filtre = 'aylik'
        sql_base += " AND DATEDIFF(day, o.odeme_tarihi, GETDATE()) <= 30 ORDER BY o.odeme_tarihi DESC"
        filtre_aciklama = "Son 30 Gün (Aylık Toplam Ciro)"
        filtrelenen_toplam = aylik_tutar

    cursor.execute(sql_base, params)
    odemeler = cursor.fetchall()
    conn.close()

    content = render_template_string(KASA_RAPORLARI_CONTENT, 
                                   gunluk_tutar=gunluk_tutar,
                                   gunluk_adet=gunluk_adet,
                                   haftalik_tutar=haftalik_tutar, 
                                   haftalik_adet=haftalik_adet,
                                   aylik_tutar=aylik_tutar, 
                                   aylik_adet=aylik_adet,
                                   odemeler=odemeler,
                                   aktif_filtre=filtre,
                                   filtre_aciklama=filtre_aciklama,
                                   filtrelenen_toplam=filtrelenen_toplam,
                                   baslangic=baslangic,
                                   bitis=bitis,
                                   grafik_etiketleri=grafik_etiketleri,
                                   grafik_verileri=grafik_verileri)
    return render_template_string(BASE_HTML, content=content)

@app.route('/kasa-raporlari/z-raporu')
def z_raporu():
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT ISNULL(SUM(tutar), 0), COUNT(*) 
        FROM Odemeler 
        WHERE durum = N'Ödendi' AND CAST(odeme_tarihi AS DATE) = CAST(GETDATE() AS DATE)
    """)
    row = cursor.fetchone()
    toplam_ciro = float(row[0] or 0.0)
    toplam_islem = row[1]

    cursor.execute("""
        SELECT odeme_yontemi, COUNT(*), SUM(tutar)
        FROM Odemeler
        WHERE durum = N'Ödendi' AND CAST(odeme_tarihi AS DATE) = CAST(GETDATE() AS DATE)
        GROUP BY odeme_yontemi
    """)
    detaylar = cursor.fetchall()
    conn.close()

    from datetime import datetime
    bugun = datetime.now().strftime('%d.%m.%Y')

    return render_template_string(Z_RAPORU_HTML, bugun=bugun, toplam_ciro=toplam_ciro, toplam_islem=toplam_islem, detaylar=detaylar)

@app.route('/kasa-raporlari/excel')
def kasa_excel_indir():
    filtre = request.args.get('filtre', 'aylik')
    baslangic = request.args.get('baslangic', '')
    bitis = request.args.get('bitis', '')

    conn = get_db()
    cursor = conn.cursor()

    sql_base = """
        SELECT o.odeme_id, o.odeme_tarihi, o.odeme_yontemi, o.tutar, o.aciklama,
               h.hasta_no, h.ad, h.soyad, h.hasta_id
        FROM Odemeler o
        INNER JOIN Hastalar h ON o.hasta_id = h.hasta_id
        WHERE o.durum = N'Ödendi'
    """
    params = []

    if filtre == 'gunluk':
        sql_base += " AND CAST(o.odeme_tarihi AS DATE) = CAST(GETDATE() AS DATE) ORDER BY o.odeme_tarihi DESC"
        baslik = "Bugunku_Tahsilat_Raporu"
    elif filtre == 'haftalik':
        sql_base += " AND DATEDIFF(day, o.odeme_tarihi, GETDATE()) <= 7 ORDER BY o.odeme_tarihi DESC"
        baslik = "Haftalik_Ciro_Raporu"
    elif filtre == 'ozel' and baslangic and bitis:
        sql_base += " AND CAST(o.odeme_tarihi AS DATE) >= ? AND CAST(o.odeme_tarihi AS DATE) <= ? ORDER BY o.odeme_tarihi DESC"
        params = [baslangic, bitis]
        baslik = f"Kasa_Raporu_{baslangic}_ile_{bitis}"
    else:
        sql_base += " AND DATEDIFF(day, o.odeme_tarihi, GETDATE()) <= 30 ORDER BY o.odeme_tarihi DESC"
        baslik = "Aylik_Ciro_Raporu"

    cursor.execute(sql_base, params)
    rows = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Kasa Raporu"

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin', color='D3D3D3'),
                         right=Side(style='thin', color='D3D3D3'),
                         top=Side(style='thin', color='D3D3D3'),
                         bottom=Side(style='thin', color='D3D3D3'))

    ws.append(["DIS HASTANESİ BİLGİ YÖNETİM SİSTEMİ - KASA RAPORU"])
    ws.merge_cells("A1:G1")
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.append([])

    headers = ["Makbuz No", "Tarih", "Hasta No", "Hasta Adı Soyadı", "Ödeme Yöntemi", "Açıklama", "Tutar (₺)"]
    ws.append(headers)
    ws.row_dimensions[3].height = 22

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    toplam_tutar = 0.0
    for row in rows:
        makbuz_no = f"MAK-{str(row[0]).zfill(5)}"
        tarih_str = str(row[1])[:16]
        h_no = row[5] or f"HST-{row[8]}"
        h_ad = f"{row[6]} {row[7]}"
        yontem = row[2]
        aciklama = row[4] or "Genel Tahsilat"
        tutar = float(row[3] or 0.0)
        toplam_tutar += tutar

        ws.append([makbuz_no, tarih_str, h_no, h_ad, yontem, aciklama, tutar])

    ws.append([])
    ws.append(["", "", "", "", "", "TOPLAM TAHSİLAT:", toplam_tutar])
    last_row = ws.max_row
    ws.cell(row=last_row, column=6).font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=last_row, column=7).font = Font(name="Calibri", size=11, bold=True, color="059669")
    ws.cell(row=last_row, column=7).number_format = '#,##0.00 "₺"'

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=f"{baslik}.xlsx")

@app.route('/kasa-raporlari/hekim-performans')
def hekim_performans_raporu():
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            dh.hekim_id,
            dh.ad,
            dh.soyad,
            b.brans_adi,
            dh.oda_no,
            COUNT(DISTINCT r.hasta_id) AS toplam_hasta,
            COUNT(ti.islem_id) AS toplam_islem,
            ISNULL(SUM(ti.birim_fiyat), 0) AS uretilen_ciro
        FROM DisHekimleri dh
        INNER JOIN Branslar b ON dh.brans_id = b.brans_id
        LEFT JOIN Randevular r ON dh.hekim_id = r.hekim_id
        LEFT JOIN Tedaviler t ON r.randevu_id = t.randevu_id
        LEFT JOIN TedaviIslemleri ti ON t.tedavi_id = ti.tedavi_id
        GROUP BY dh.hekim_id, dh.ad, dh.soyad, b.brans_adi, dh.oda_no
        ORDER BY uretilen_ciro DESC, toplam_hasta DESC
    """)
    raw_performans = cursor.fetchall()

    toplam_ciro = sum(float(hp[7] or 0) for hp in raw_performans)
    
    hekim_performans = []
    for hp in raw_performans:
        h_id, ad, soyad, brans, oda, h_sayisi, i_sayisi, ciro = hp
        c_float = float(ciro or 0)
        yuzde = round((c_float / toplam_ciro * 100), 1) if toplam_ciro > 0 else 0.0
        hekim_performans.append((h_id, ad, soyad, brans, oda, h_sayisi, i_sayisi, c_float, yuzde))

    lider_hekim = f"Dt. {hekim_performans[0][1]} {hekim_performans[0][2]}" if hekim_performans and hekim_performans[0][7] > 0 else "Henüz Kayıt Yok"

    conn.close()

    content = render_template_string(HEKIM_PERFORMANS_CONTENT,
                                   hekim_performans=hekim_performans,
                                   toplam_ciro=toplam_ciro,
                                   lider_hekim=lider_hekim)
    return render_template_string(BASE_HTML, content=content)

@app.route('/odeme-makbuz/<int:odeme_id>')
def odeme_makbuz(odeme_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT o.odeme_id, CONVERT(VARCHAR(16), o.odeme_tarihi, 120), o.odeme_yontemi, o.tutar,
               h.hasta_id, h.ad, h.soyad, h.hasta_no, o.aciklama
        FROM Odemeler o
        INNER JOIN Hastalar h ON o.hasta_id = h.hasta_id
        WHERE o.odeme_id = ?
    """, (odeme_id,))
    odeme = cursor.fetchone()
    conn.close()

    if not odeme:
        flash('Ödeme belgesi bulunamadı.', 'warning')
        return redirect(url_for('kasa_raporlari'))

    return render_template_string(MAKBUZ_HTML, odeme=odeme)

@app.route('/stoklar')
def stok_listesi():
    arama = request.args.get('q', '').strip()
    conn = get_db()
    cursor = conn.cursor()
    
    if arama:
        cursor.execute("""
            SELECT malzeme_id, malzeme_adi, kategori, miktar, birim, kritik_seviye,
                   CONVERT(VARCHAR(16), son_guncelleme, 120) AS guncelleme
            FROM Stoklar
            WHERE malzeme_adi LIKE ? OR kategori LIKE ?
            ORDER BY (CASE WHEN miktar <= kritik_seviye THEN 0 ELSE 1 END), kategori ASC, malzeme_adi ASC
        """, (f"%{arama}%", f"%{arama}%"))
    else:
        cursor.execute("""
            SELECT malzeme_id, malzeme_adi, kategori, miktar, birim, kritik_seviye,
                   CONVERT(VARCHAR(16), son_guncelleme, 120) AS guncelleme
            FROM Stoklar
            ORDER BY (CASE WHEN miktar <= kritik_seviye THEN 0 ELSE 1 END), kategori ASC, malzeme_adi ASC
        """)
        
    stoklar = cursor.fetchall()
    conn.close()
    return render_template_string(BASE_HTML, content=render_template_string(STOKLAR_CONTENT, stoklar=stoklar, arama=arama))

@app.route('/stok-ekle', methods=['POST'])
def stok_ekle():
    malzeme_adi = request.form['malzeme_adi'].strip()
    kategori = request.form['kategori']
    miktar = int(request.form.get('miktar', 0))
    birim = request.form['birim']
    kritik_seviye = int(request.form.get('kritik_seviye', 10))

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT COUNT(*) FROM Stoklar WHERE malzeme_adi = ?", (malzeme_adi,))
        if cursor.fetchone()[0] > 0:
            flash('Bu malzeme zaten envanterde kayıtlı!', 'warning')
            return redirect(url_for('stok_listesi'))

        cursor.execute("""
            INSERT INTO Stoklar (malzeme_adi, kategori, miktar, birim, kritik_seviye, son_guncelleme)
            VALUES (?, ?, ?, ?, ?, GETDATE())
        """, (malzeme_adi, kategori, miktar, birim, kritik_seviye))
        conn.commit()
        flash(f"'{malzeme_adi}' ({miktar} {birim}) başarıyla stoğa eklendi.", 'success')
    except Exception as e:
        conn.rollback()
        flash(f'Stok eklenirken hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('stok_listesi'))

@app.route('/stok-manuel-guncelle', methods=['POST'])
def stok_manuel_guncelle():
    malzeme_id = request.form['malzeme_id']
    yeni_miktar = int(request.form.get('yeni_miktar', 0))

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE Stoklar SET miktar = ?, son_guncelleme = GETDATE() WHERE malzeme_id = ?", (yeni_miktar, malzeme_id))
        conn.commit()
        flash('Malzeme stok miktarı güncellendi.', 'success')
    except Exception as e:
        conn.rollback()
        flash(f'Hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('stok_listesi'))

@app.route('/stok-hareket/<int:malzeme_id>/<string:islem>')
def stok_hareket(malzeme_id, islem):
    conn = get_db()
    cursor = conn.cursor()
    try:
        if islem == 'arttir':
            cursor.execute("UPDATE Stoklar SET miktar = miktar + 1, son_guncelleme = GETDATE() WHERE malzeme_id = ?", (malzeme_id,))
            flash('Malzeme stoğu +1 artırıldı.', 'success')
        elif islem == 'azalt':
            cursor.execute("SELECT miktar, malzeme_adi FROM Stoklar WHERE malzeme_id = ?", (malzeme_id,))
            item = cursor.fetchone()
            if item and item[0] > 0:
                cursor.execute("UPDATE Stoklar SET miktar = miktar - 1, son_guncelleme = GETDATE() WHERE malzeme_id = ?", (malzeme_id,))
                flash(f"1 adet {item[1]} kullanıldı (stok düşüldü).", 'info')
            else:
                flash('Stok miktarı zaten 0! Daha fazla düşülemez.', 'danger')
        conn.commit()
    except Exception as e:
        conn.rollback()
        flash(f'Stok güncellenirken hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('stok_listesi'))

@app.route('/stok-sil/<int:malzeme_id>')
def stok_sil(malzeme_id):
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Stoklar WHERE malzeme_id = ?", (malzeme_id,))
        conn.commit()
        flash('Malzeme envanterden silindi.', 'info')
    except Exception as e:
        conn.rollback()
        flash(f'Hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('stok_listesi'))

@app.route('/hastalar')
def tum_hastalar():
    arama = request.args.get('q', '').strip()
    conn = get_db()
    cursor = conn.cursor()
    
    if arama:
        cursor.execute("""
            SELECT hasta_id, tc_no, ad, soyad, CONVERT(VARCHAR(10), dogum_tarihi, 120), cinsiyet, telefon, email, tibbi_not, hasta_no 
            FROM Hastalar 
            WHERE ad LIKE ? OR soyad LIKE ? OR tc_no LIKE ? OR tibbi_not LIKE ? OR hasta_no LIKE ?
            ORDER BY ad ASC, soyad ASC
        """, (f"%{arama}%", f"%{arama}%", f"%{arama}%", f"%{arama}%", f"%{arama}%"))
    else:
        cursor.execute("""
            SELECT hasta_id, tc_no, ad, soyad, CONVERT(VARCHAR(10), dogum_tarihi, 120), cinsiyet, telefon, email, tibbi_not, hasta_no 
            FROM Hastalar 
            ORDER BY ad ASC, soyad ASC
        """)
        
    hastalar = cursor.fetchall()
    conn.close()

    content = render_template_string(HASTALAR_CONTENT, hastalar=hastalar, arama=arama)
    return render_template_string(BASE_HTML, content=content)

@app.route('/hasta-ekle', methods=['GET', 'POST'])
def hasta_ekle():
    if request.method == 'POST':
        hasta_no = request.form.get('hasta_no', '').strip() or None
        tc_no = request.form.get('tc_no', '').strip() or None
        ad = request.form['ad'].strip()
        soyad = request.form['soyad'].strip()
        tibbi_not = request.form.get('tibbi_not', '').strip() or None
        dogum_tarihi = request.form.get('dogum_tarihi', '').strip() or None
        cinsiyet = request.form.get('cinsiyet', 'Belirtilmemiş')
        telefon = request.form.get('telefon', '').strip() or None
        email = request.form.get('email', '').strip() or None
        adres = request.form.get('adres', '').strip() or None

        conn = get_db()
        cursor = conn.cursor()
        try:
            if tc_no:
                cursor.execute("SELECT COUNT(*) FROM Hastalar WHERE tc_no = ?", (tc_no,))
                if cursor.fetchone()[0] > 0:
                    flash('Bu T.C. Kimlik Numarası zaten kayıtlı!', 'warning')
                    return redirect(url_for('hasta_ekle'))

            cursor.execute("""
                INSERT INTO Hastalar (tc_no, ad, soyad, dogum_tarihi, cinsiyet, telefon, email, adres, tibbi_not, hasta_no)
                OUTPUT INSERTED.hasta_id
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (tc_no, ad, soyad, dogum_tarihi, cinsiyet, telefon, email, adres, tibbi_not, hasta_no))
            
            yeni_id = cursor.fetchone()[0]
            if not hasta_no:
                otomatik_no = f"HST-{str(yeni_id).zfill(4)}"
                cursor.execute("UPDATE Hastalar SET hasta_no = ? WHERE hasta_id = ?", (otomatik_no, yeni_id))

            conn.commit()
            flash(f'{ad} {soyad} isimli hasta başarıyla kaydedildi.', 'success')
            return redirect(url_for('tum_hastalar'))
        except Exception as e:
            conn.rollback()
            flash(f'Hasta kaydedilirken hata oluştu: {e}', 'danger')
            return redirect(url_for('hasta_ekle'))
        finally:
            conn.close()

    return render_template_string(BASE_HTML, content=render_template_string(HASTA_EKLE_CONTENT))

@app.route('/hasta/<int:hasta_id>/duzenle', methods=['GET', 'POST'])
def hasta_duzenle(hasta_id):
    conn = get_db()
    cursor = conn.cursor()
    
    if request.method == 'POST':
        hasta_no = request.form.get('hasta_no', '').strip() or None
        tc_no = request.form.get('tc_no', '').strip() or None
        ad = request.form['ad'].strip()
        soyad = request.form['soyad'].strip()
        tibbi_not = request.form.get('tibbi_not', '').strip() or None
        dogum_tarihi = request.form.get('dogum_tarihi', '').strip() or None
        cinsiyet = request.form.get('cinsiyet', 'Belirtilmemiş')
        telefon = request.form.get('telefon', '').strip() or None
        email = request.form.get('email', '').strip() or None
        adres = request.form.get('adres', '').strip() or None

        try:
            cursor.execute("""
                UPDATE Hastalar 
                SET hasta_no = ?, tc_no = ?, ad = ?, soyad = ?, dogum_tarihi = ?, cinsiyet = ?, telefon = ?, email = ?, adres = ?, tibbi_not = ?
                WHERE hasta_id = ?
            """, (hasta_no, tc_no, ad, soyad, dogum_tarihi, cinsiyet, telefon, email, adres, tibbi_not, hasta_id))
            conn.commit()
            flash(f'{ad} {soyad} isimli hastanın bilgileri başarıyla güncellendi.', 'success')
            return redirect(url_for('hasta_detay', hasta_id=hasta_id))
        except Exception as e:
            conn.rollback()
            flash(f'Bilgiler güncellenirken hata: {e}', 'danger')
            return redirect(url_for('hasta_duzenle', hasta_id=hasta_id))
        finally:
            conn.close()

    cursor.execute("""
        SELECT hasta_id, tc_no, ad, soyad, CONVERT(VARCHAR(10), dogum_tarihi, 120), cinsiyet, telefon, email, tibbi_not, hasta_no, adres 
        FROM Hastalar WHERE hasta_id = ?
    """, (hasta_id,))
    hasta = cursor.fetchone()
    conn.close()
    
    return render_template_string(BASE_HTML, content=render_template_string(HASTA_DUZENLE_CONTENT, hasta=hasta))

@app.route('/hekim-ekle', methods=['GET', 'POST'])
def hekim_ekle():
    conn = get_db()
    cursor = conn.cursor()
    if request.method == 'POST':
        ad = request.form['ad'].strip()
        soyad = request.form['soyad'].strip()
        brans_id = request.form['brans_id']
        oda_no = request.form['oda_no'].strip()
        telefon = request.form['telefon'].strip()
        email = request.form['email'].strip()

        try:
            cursor.execute("SELECT COUNT(*) FROM DisHekimleri WHERE email = ?", (email,))
            if cursor.fetchone()[0] > 0:
                flash('Bu e-posta adresiyle kayıtlı bir hekim zaten var!', 'warning')
                return redirect(url_for('hekim_ekle'))

            cursor.execute("""
                INSERT INTO DisHekimleri (brans_id, ad, soyad, telefon, email, oda_no)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (brans_id, ad, soyad, telefon, email, oda_no))
            conn.commit()
            flash(f'Dt. {ad} {soyad} hekim kadrosuna eklendi.', 'success')
            return redirect(url_for('index'))
        except Exception as e:
            conn.rollback()
            flash(f'Hekim eklenirken hata: {e}', 'danger')
            return redirect(url_for('hekim_ekle'))
        finally:
            conn.close()

    cursor.execute("SELECT brans_id, brans_adi FROM Branslar ORDER BY brans_adi ASC")
    branslar = cursor.fetchall()
    conn.close()
    return render_template_string(BASE_HTML, content=render_template_string(HEKIM_EKLE_CONTENT, branslar=branslar))

@app.route('/hasta/<int:hasta_id>')
def hasta_detay(hasta_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT hasta_id, tc_no, ad, soyad, CONVERT(VARCHAR(10), dogum_tarihi, 120), cinsiyet, telefon, email, tibbi_not, hasta_no 
        FROM Hastalar WHERE hasta_id = ?
    """, (hasta_id,))
    hasta = cursor.fetchone()
    
    cursor.execute("""
        SELECT tp.plan_id, tp.hasta_id, tp.tahmini_tutar, tp.plan_adi, tp.aciklama,
               tp.toplam_seans, tp.tamamlanan_seans, tp.durum,
               CONVERT(VARCHAR(10), tp.olusturma_tarihi, 120), tp.hekim_id,
               (dh.ad + ' ' + dh.soyad) AS hekim_ad
        FROM TedaviPlanlari tp
        INNER JOIN DisHekimleri dh ON tp.hekim_id = dh.hekim_id
        WHERE tp.hasta_id = ?
        ORDER BY tp.olusturma_tarihi DESC
    """, (hasta_id,))
    planlar = cursor.fetchall()

    cursor.execute("""
        SELECT t.tedavi_id, t.tani, CONVERT(VARCHAR(10), t.baslangic_tarihi, 120), ti.dis_no, di.islem_adi, ti.birim_fiyat
        FROM Tedaviler t
        INNER JOIN Randevular r ON t.randevu_id = r.randevu_id
        LEFT JOIN TedaviIslemleri ti ON t.tedavi_id = ti.tedavi_id
        LEFT JOIN DisIslemleri di ON ti.islem_id = di.islem_id
        WHERE r.hasta_id = ?
        ORDER BY t.baslangic_tarihi DESC
    """, (hasta_id,))
    tedaviler = cursor.fetchall()

    cursor.execute("""
        SELECT tutar, CONVERT(VARCHAR(10), odeme_tarihi, 120), odeme_yontemi, durum, aciklama, odeme_id
        FROM Odemeler WHERE hasta_id = ? ORDER BY odeme_tarihi DESC
    """, (hasta_id,))
    odemeler = cursor.fetchall()

    cursor.execute("""
        SELECT ISNULL(SUM(ti.adet * ti.birim_fiyat), 0)
        FROM Tedaviler t
        INNER JOIN Randevular r ON t.randevu_id = r.randevu_id
        INNER JOIN TedaviIslemleri ti ON t.tedavi_id = ti.tedavi_id
        WHERE r.hasta_id = ?
    """, (hasta_id,))
    toplam_tedavi_tutari = float(cursor.fetchone()[0] or 0.0)

    cursor.execute("""
        SELECT ISNULL(SUM(tutar), 0)
        FROM Odemeler
        WHERE hasta_id = ? AND durum = N'Ödendi'
    """, (hasta_id,))
    toplam_odenen = float(cursor.fetchone()[0] or 0.0)

    kalan_borc = max(0.0, toplam_tedavi_tutari - toplam_odenen)

    cursor.execute("SELECT islem_id, islem_adi FROM DisIslemleri")
    dis_islemleri = cursor.fetchall()

    cursor.execute("SELECT hekim_id, ad, soyad FROM DisHekimleri")
    tum_hekimler = cursor.fetchall()

    conn.close()

    content = render_template_string(HASTA_DETAY_CONTENT, 
                                   hasta=hasta, 
                                   planlar=planlar,
                                   tedaviler=tedaviler, 
                                   odemeler=odemeler,
                                   toplam_tedavi_tutari=toplam_tedavi_tutari,
                                   toplam_odenen=toplam_odenen,
                                   kalan_borc=kalan_borc,
                                   dis_islemleri=dis_islemleri, 
                                   tum_hekimler=tum_hekimler)
    return render_template_string(BASE_HTML, content=content)

@app.route('/hasta/<int:hasta_id>/plan-ekle', methods=['POST'])
def plan_ekle(hasta_id):
    plan_adi = request.form['plan_adi'].strip()
    hekim_id = request.form['hekim_id']
    toplam_seans = int(request.form.get('toplam_seans', 1))
    tahmini_tutar = float(request.form.get('tahmini_tutar', 0.0))
    aciklama = request.form.get('aciklama', '').strip()

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO TedaviPlanlari (hasta_id, hekim_id, plan_adi, aciklama, toplam_seans, tamamlanan_seans, tahmini_tutar, durum, olusturma_tarihi)
            VALUES (?, ?, ?, ?, ?, 0, ?, N'Devam Ediyor', GETDATE())
        """, (hasta_id, hekim_id, plan_adi, aciklama, toplam_seans, tahmini_tutar))
        conn.commit()
        flash(f"'{plan_adi}' tedavi planı ({toplam_seans} Seans) ve fiyat teklifi oluşturuldu.", 'success')
    except Exception as e:
        conn.rollback()
        flash(f'Plan eklenirken hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('hasta_detay', hasta_id=hasta_id))

@app.route('/plan-seans-arttir/<int:plan_id>')
def plan_seans_arttir(plan_id):
    hasta_id = request.args.get('hasta_id')
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT tamamlanan_seans, toplam_seans FROM TedaviPlanlari WHERE plan_id = ?", (plan_id,))
        row = cursor.fetchone()
        if row:
            yeni_seans = row[0] + 1
            yeni_durum = 'Tamamlandı' if yeni_seans >= row[1] else 'Devam Ediyor'
            cursor.execute("""
                UPDATE TedaviPlanlari 
                SET tamamlanan_seans = ?, durum = ?
                WHERE plan_id = ?
            """, (yeni_seans, yeni_durum, plan_id))
            conn.commit()
            flash(f'Tedavi planında seans ilerletildi ({yeni_seans}/{row[1]}). Durum: {yeni_durum}', 'success')
    except Exception as e:
        conn.rollback()
        flash(f'Hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('hasta_detay', hasta_id=hasta_id))

@app.route('/plan-sil/<int:plan_id>')
def plan_sil(plan_id):
    hasta_id = request.args.get('hasta_id')
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM TedaviPlanlari WHERE plan_id = ?", (plan_id,))
        conn.commit()
        flash('Tedavi planı silindi.', 'info')
    except Exception as e:
        conn.rollback()
        flash(f'Hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('hasta_detay', hasta_id=hasta_id))

@app.route('/plan-teklif-yazdir/<int:plan_id>')
def plan_teklif_yazdir(plan_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT tp.plan_id, tp.hasta_id, tp.tahmini_tutar, tp.plan_adi, tp.aciklama,
               tp.toplam_seans, tp.tamamlanan_seans, tp.durum,
               CONVERT(VARCHAR(10), tp.olusturma_tarihi, 120),
               h.ad, h.soyad, h.hasta_no,
               dh.ad, dh.soyad
        FROM TedaviPlanlari tp
        INNER JOIN Hastalar h ON tp.hasta_id = h.hasta_id
        INNER JOIN DisHekimleri dh ON tp.hekim_id = dh.hekim_id
        WHERE tp.plan_id = ?
    """, (plan_id,))
    plan = cursor.fetchone()
    conn.close()

    if not plan:
        flash('Plan kaydı bulunamadı.', 'warning')
        return redirect(url_for('tum_hastalar'))

    return render_template_string(PLAN_TEKLIF_HTML, plan=plan)

@app.route('/hasta/<int:hasta_id>/gecmis')
def hasta_gecmis(hasta_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT hasta_id, tc_no, ad, soyad, CONVERT(VARCHAR(10), dogum_tarihi, 120), cinsiyet, telefon, email, tibbi_not, hasta_no 
        FROM Hastalar WHERE hasta_id = ?
    """, (hasta_id,))
    hasta = cursor.fetchone()

    cursor.execute("""
        SELECT CONVERT(VARCHAR(10), r.randevu_tarihi, 120), CONVERT(VARCHAR(5), r.randevu_saati, 108),
               (dh.ad + ' ' + dh.soyad), r.durum, r.notlar
        FROM Randevular r
        INNER JOIN DisHekimleri dh ON r.hekim_id = dh.hekim_id
        WHERE r.hasta_id = ?
        ORDER BY r.randevu_tarihi DESC
    """, (hasta_id,))
    randevular = cursor.fetchall()

    cursor.execute("""
        SELECT CONVERT(VARCHAR(10), t.baslangic_tarihi, 120), t.tani, ti.dis_no, di.islem_adi, ti.birim_fiyat
        FROM Tedaviler t
        INNER JOIN Randevular r ON t.randevu_id = r.randevu_id
        INNER JOIN TedaviIslemleri ti ON t.tedavi_id = ti.tedavi_id
        INNER JOIN DisIslemleri di ON ti.islem_id = di.islem_id
        WHERE r.hasta_id = ?
        ORDER BY t.baslangic_tarihi DESC
    """, (hasta_id,))
    tedaviler = cursor.fetchall()

    conn.close()
    content = render_template_string(HASTA_GECMIS_CONTENT, hasta=hasta, randevular=randevular, tedaviler=tedaviler)
    return render_template_string(BASE_HTML, content=content)

@app.route('/hasta-sil/<int:hasta_id>')
def hasta_sil(hasta_id):
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM TedaviPlanlari WHERE hasta_id = ?", (hasta_id,))
        cursor.execute("DELETE FROM Odemeler WHERE hasta_id = ?", (hasta_id,))
        cursor.execute("DELETE FROM HastaSigorta WHERE hasta_id = ?", (hasta_id,))
        
        cursor.execute("""
            DELETE FROM TedaviIslemleri 
            WHERE tedavi_id IN (SELECT t.tedavi_id FROM Tedaviler t INNER JOIN Randevular r ON t.randevu_id = r.randevu_id WHERE r.hasta_id = ?)
        """, (hasta_id,))
        
        cursor.execute("""
            DELETE FROM Tedaviler 
            WHERE randevu_id IN (SELECT randevu_id FROM Randevular WHERE hasta_id = ?)
        """, (hasta_id,))
        
        cursor.execute("DELETE FROM Randevular WHERE hasta_id = ?", (hasta_id,))
        cursor.execute("DELETE FROM Hastalar WHERE hasta_id = ?", (hasta_id,))
        
        conn.commit()
        flash('Hasta ve tüm tıbbi geçmişi silindi.', 'info')
    except Exception as e:
        conn.rollback()
        flash(f'Hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('tum_hastalar'))

@app.route('/randevu-sil/<int:randevu_id>')
def randevu_sil(randevu_id):
    hekim_id = request.args.get('hekim_id')
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM TedaviIslemleri WHERE tedavi_id IN (SELECT tedavi_id FROM Tedaviler WHERE randevu_id = ?)", (randevu_id,))
        cursor.execute("DELETE FROM Tedaviler WHERE randevu_id = ?", (randevu_id,))
        cursor.execute("DELETE FROM Randevular WHERE randevu_id = ?", (randevu_id,))
        conn.commit()
        flash('Randevu kaydı silindi.', 'info')
    except Exception as e:
        conn.rollback()
        flash(f'Hata: {e}', 'danger')
    finally:
        conn.close()

    if hekim_id:
        return redirect(url_for('hekim_randevulari', hekim_id=hekim_id))
    return redirect(url_for('index'))

@app.route('/tedavi-sil/<int:tedavi_id>')
def tedavi_sil(tedavi_id):
    hasta_id = request.args.get('hasta_id')
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM TedaviIslemleri WHERE tedavi_id = ?", (tedavi_id,))
        cursor.execute("DELETE FROM Tedaviler WHERE tedavi_id = ?", (tedavi_id,))
        conn.commit()
        flash('Tedavi kaydı başarıyla silindi.', 'info')
    except Exception as e:
        conn.rollback()
        flash(f'Hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('hasta_detay', hasta_id=hasta_id))

@app.route('/hekim/<int:hekim_id>/randevular')
def hekim_randevulari(hekim_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT hekim_id, ad, soyad, email FROM DisHekimleri WHERE hekim_id = ?", (hekim_id,))
    hekim = cursor.fetchone()
    
    cursor.execute("""
        SELECT r.randevu_id, h.hasta_id, (h.ad + ' ' + h.soyad) AS hasta_ad, h.telefon, 
               CONVERT(VARCHAR(10), r.randevu_tarihi, 120), CONVERT(VARCHAR(5), r.randevu_saati, 108) AS saat, r.durum, r.notlar
        FROM Randevular r
        INNER JOIN Hastalar h ON r.hasta_id = h.hasta_id
        WHERE r.hekim_id = ?
        ORDER BY r.randevu_tarihi DESC, r.randevu_saati ASC
    """, (hekim_id,))
    randevular = cursor.fetchall()
    conn.close()

    content = render_template_string(RANDEVULAR_CONTENT, hekim=hekim, randevular=randevular)
    return render_template_string(BASE_HTML, content=content)

@app.route('/hekim/<int:hekim_id>/bugun-ozet-gonder')
def bugun_ozet_gonder(hekim_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT ad, soyad, email FROM DisHekimleri WHERE hekim_id = ?", (hekim_id,))
    hekim = cursor.fetchone()

    cursor.execute("""
        SELECT CONVERT(VARCHAR(5), r.randevu_saati, 108), h.ad + ' ' + h.soyad
        FROM Randevular r
        INNER JOIN Hastalar h ON r.hasta_id = h.hasta_id
        WHERE r.hekim_id = ? AND r.randevu_tarihi = CAST(GETDATE() AS DATE)
        ORDER BY r.randevu_saati ASC
    """, (hekim_id,))
    bugunku_randevular = cursor.fetchall()
    conn.close()

    if not bugunku_randevular:
        flash(f"Dt. {hekim[0]} {hekim[1]} için bugün (aktif tarihte) kayıtlı randevu bulunmuyor.", "warning")
        return redirect(url_for('hekim_randevulari', hekim_id=hekim_id))

    randevu_metni = ", ".join([f"{r[0]} - {r[1]}" for r in bugunku_randevular])
    flash(f"Başarılı! Dt. {hekim[0]} {hekim[1]} ({hekim[2]}) hekimine bugünkü tüm randevuları ({randevu_metni}) içeren günlük hatırlatma özeti iletildi.", "success")
    return redirect(url_for('hekim_randevulari', hekim_id=hekim_id))

@app.route('/hekim-mail-hatirlat', methods=['POST'])
def hekim_mail_hatirlat():
    hekim_id = request.form['hekim_id']
    hekim_mail = request.form['hekim_mail']
    hatirlatma_saati = request.form['hatirlatma_saati']
    
    flash(f"Hekime ({hekim_mail}) saat {hatirlatma_saati} için randevu hatırlatma bildirimi planlandı.", 'success')
    return redirect(url_for('hekim_randevulari', hekim_id=hekim_id))

@app.route('/hasta/<int:hasta_id>/tedavi-ekle', methods=['POST'])
def tedavi_ekle(hasta_id):
    tani = request.form['tani']
    islem_id = request.form['islem_id']
    dis_no = request.form.get('dis_no', '')
    fiyat = request.form['fiyat']

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT randevu_id FROM Randevular WHERE hasta_id = ? ORDER BY randevu_id DESC", (hasta_id,))
        randevu = cursor.fetchone()
        
        if randevu:
            randevu_id = randevu[0]
        else:
            cursor.execute("SELECT TOP 1 hekim_id FROM DisHekimleri ORDER BY hekim_id ASC")
            hekim_id = cursor.fetchone()[0]
            cursor.execute("""
                INSERT INTO Randevular (hasta_id, hekim_id, randevu_tarihi, randevu_saati, durum, notlar)
                OUTPUT INSERTED.randevu_id
                VALUES (?, ?, CAST(GETDATE() AS DATE), '09:00', N'Tamamlandı', N'Klinik Tedavi Girişi')
            """, (hasta_id, hekim_id))
            randevu_id = cursor.fetchone()[0]

        cursor.execute("""
            INSERT INTO Tedaviler (randevu_id, tani, baslangic_tarihi)
            OUTPUT INSERTED.tedavi_id
            VALUES (?, ?, CAST(GETDATE() AS DATE))
        """, (randevu_id, tani))
        tedavi_id = cursor.fetchone()[0]

        cursor.execute("""
            INSERT INTO TedaviIslemleri (tedavi_id, islem_id, dis_no, birim_fiyat, uygulama_tarihi)
            VALUES (?, ?, ?, ?, CAST(GETDATE() AS DATE))
        """, (tedavi_id, islem_id, dis_no, fiyat))

        conn.commit()
        flash(f'{dis_no if dis_no else "Genel"} no\'lu işlem/tedavi kaydı başarıyla oluşturuldu.', 'success')
    except Exception as e:
        conn.rollback()
        flash(f'Tedavi kaydedilirken hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('hasta_detay', hasta_id=hasta_id))

@app.route('/hasta/<int:hasta_id>/odeme-ekle', methods=['POST'])
def odeme_ekle(hasta_id):
    tutar = request.form['tutar']
    odeme_yontemi = request.form['odeme_yontemi']
    durum = request.form['durum']
    aciklama = request.form.get('aciklama', '')

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO Odemeler (hasta_id, tutar, odeme_tarihi, odeme_yontemi, durum, aciklama)
            OUTPUT INSERTED.odeme_id
            VALUES (?, ?, GETDATE(), ?, ?, ?)
        """, (hasta_id, tutar, odeme_yontemi, durum, aciklama))
        odeme_id = cursor.fetchone()[0]
        conn.commit()
        flash(f'Ödeme tahsilatı (#MAK-{str(odeme_id).zfill(5)}) başarıyla işlendi ve borçtan düşüldü.', 'success')
    except Exception as e:
        conn.rollback()
        flash(f'Hata: {e}', 'danger')
    finally:
        conn.close()

    return redirect(url_for('hasta_detay', hasta_id=hasta_id))

@app.route('/randevu-ekle', methods=['GET', 'POST'])
def randevu_ekle():
    if request.method == 'POST':
        hasta_id = request.form['hasta_id']
        hekim_id = request.form['hekim_id']
        tarih = request.form['tarih']
        saat = request.form['saat']
        notlar = request.form.get('notlar', '')

        conn = get_db()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                SELECT COUNT(*) FROM Randevular 
                WHERE hekim_id = ? AND randevu_tarihi = ? AND randevu_saati = ?
            """, (hekim_id, tarih, saat))
            
            if cursor.fetchone()[0] > 0:
                flash('Hekimin seçilen saatte randevusu mevcut.', 'warning')
                return redirect(url_for('randevu_ekle'))

            cursor.execute("""
                INSERT INTO Randevular (hasta_id, hekim_id, randevu_tarihi, randevu_saati, durum, notlar)
                VALUES (?, ?, ?, ?, N'Bekliyor', ?)
            """, (hasta_id, hekim_id, tarih, saat, notlar))
            
            conn.commit()
            flash('Randevu kaydı başarıyla oluşturuldu.', 'success')
            return redirect(url_for('hekim_randevulari', hekim_id=hekim_id))
        except Exception as e:
            conn.rollback()
            flash(f'Randevu ekleme hatası: {e}', 'danger')
            return redirect(url_for('randevu_ekle'))
        finally:
            conn.close()

    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT hasta_id, ad, soyad, tc_no, tibbi_not, hasta_no FROM Hastalar ORDER BY ad, soyad")
        hastalar = cursor.fetchall()
        
        cursor.execute("SELECT hekim_id, ad, soyad FROM DisHekimleri ORDER BY ad, soyad")
        hekimler = cursor.fetchall()
    finally:
        conn.close()

    content = render_template_string(YENI_RANDEVU_CONTENT, hastalar=hastalar, hekimler=hekimler)
    return render_template_string(BASE_HTML, content=content)

if __name__ == '__main__':
    app.run(debug=True)
