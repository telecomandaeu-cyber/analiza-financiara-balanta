# Platformă de Analiză a Poziției Financiare și a Performanțelor Financiare (.DOCX)

Aplicație automatizată în Python care extrage datele contabile din orice **Balanță de verificare (`Balanta.pdf`)**, calculează cei **20 de piloni ai diagnosticului economico-financiar** pe cele trei perioade de raportare (`Perioada de raportare 0`, `Perioada de raportare 1` și `Perioada de raportare 2`) și exportă raportul complet în format **Microsoft Word (`.docx`)**.

---

## Cum poate fi rulat direct din GitHub (4 moduri)

### 1. Site Web Public Gratuit 24/7 prin Streamlit Community Cloud (Recomandat)
Poți transforma acest repository GitHub într-un site web accesibil de oriunde, gratuit:
1. Intră pe [share.streamlit.io](https://share.streamlit.io) și autentifică-te cu contul tău de GitHub.
2. Apasă pe **Create app** -> selectează acest repository GitHub, ramura `main`, iar la **Main file path** scrie:
   ```text
   streamlit_app.py
   ```
3. Apasă **Deploy**. În ~1 minut primești un link public permanent unde oricine poate încărca un fișier `Balanta.pdf` și descărca raportul `.docx`.

---

### 2. Rulare cu un singur clic în browser prin GitHub Codespaces
Proiectul conține configurația `.devcontainer/devcontainer.json`:
1. Pe pagina repository-ului din GitHub, apasă pe butonul verde **`<> Code`** -> tab-ul **Codespaces** -> **Create codespace on main**.
2. Mediul instalează automat dependențele din `requirements.txt`, pornește `site_analiza_financiara.py` pe portul `8080` și deschide automat interfața web în browser.

---

### 3. Generare automată în GitHub Actions (fără niciun server)
Proiectul include fluxul automat `.github/workflows/genereaza_raport.yml`:
1. În pagina GitHub a proiectului, mergi la tab-ul **Actions** -> **Generare Raport Analiza Financiara Word (.DOCX)**.
2. Apasă pe **Run workflow**.
3. La finalul rulării (~20 secunde), descarcă documentul Word generat direct din secțiunea **Artifacts** (`Raport-Analiza-Financiara-Word`).

---

### 4. Rulare pe calculatorul propriu sau în Rețeaua Locală (LAN / Wi-Fi)
```bash
git clone <URL_REPOSITORY_GITHUB>
cd Analiza
pip install -r requirements.txt
python site_analiza_financiara.py
```
* Pe Windows, poți da direct dublu-clic pe **`Porneste_Site_Analiza.bat`** (și o singură dată pe **`Permite_Acces_Retea_Firewall.bat`** pentru a permite accesul colegilor din rețeaua locală la `http://<IP_LOCAL>:8080`).

