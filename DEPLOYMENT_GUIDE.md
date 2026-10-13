# 🚀 Gola PDF: Production Deployment Guide ($14/mo VPS)

This guide walks you through deploying **Gola PDF** to a live production server on a **\$7 to \$14/month Linux VPS** with high speed, free SSL encryption (HTTPS), and zero permanent storage retention.

---

## 1. 🖥️ Recommended VPS Providers ($6 – $14/month)

Choose any of these reliable VPS providers (select **Ubuntu 22.04 LTS** or **Ubuntu 24.04 LTS**):

| Provider | Plan | Specs | Price | Recommendation |
| :--- | :--- | :--- | :--- | :--- |
| **Hetzner Cloud** | CPX21 | 3 vCPU, 4 GB RAM, 80 GB NVMe | ~\$8.50 / mo | 🌟 **Top Pick** (Fastest NVMe, lowest latency) |
| **Contabo** | Cloud VPS 1 | 4 vCPU, 6 GB RAM, 100 GB NVMe | ~\$6.00 / mo | 💰 **Best Budget & RAM** |
| **DigitalOcean** | Basic Droplet | 1 vCPU, 2 GB RAM, 50 GB SSD | \$12.00 / mo | 🛠️ **Easiest Dashboard & Backups** |
| **Hostinger** | KVM 2 | 2 vCPU, 4 GB RAM, 50 GB NVMe | ~\$7.99 / mo | ⚡ **Great global data centers** |

---

## 2. 🌐 Pointing Your Domain DNS

Before running the deployment script, connect your domain name (e.g., `golapdf.com`) to your VPS IP address in your DNS registrar (Namecheap, GoDaddy, Cloudflare, Google Domains, etc.):

1. Add an **A Record**:
   - **Host / Name**: `@`
   - **Value / Points to**: `<YOUR_VPS_IP_ADDRESS>` (e.g., `123.45.67.89`)
   - **TTL**: `Automatic` or `300`
2. Add a second **A Record** (or CNAME) for `www`:
   - **Host / Name**: `www`
   - **Value / Points to**: `<YOUR_VPS_IP_ADDRESS>`

*(DNS propagation usually takes 5 to 15 minutes).*

---

## 3. ⚡ 1-Click Automated Deployment

Connect to your VPS via SSH from your terminal or PowerShell:
```bash
ssh root@<YOUR_VPS_IP_ADDRESS>
```

Clone your repository (or copy your files using `git` or `scp`):
```bash
git clone https://github.com/yourusername/gola_pdf.git /var/www/golapdf
cd /var/www/golapdf
```

Run the automated deployment script:
```bash
sudo bash deploy.sh golapdf.com your_email@example.com
```

### What `deploy.sh` Does Automatically:
1. Installs all required Linux packages (`Python 3.11`, `Tesseract OCR`, `Ghostscript`, `Nginx`, `Certbot`, `UFW firewall`).
2. Configures a secured Python virtual environment with all dependencies.
3. Configures a `systemd` background daemon (`golapdf.service`) with auto-restart on crashes and auto-start on boot.
4. Configures Nginx with GZIP compression, HTTP security headers, and reverse proxying.
5. Issues a free **Let's Encrypt SSL certificate** and turns on automatic HTTPS redirection!
6. Verifies local health checks and starts serving traffic.

---

## 4. 🐳 Alternative: Deploy with Docker Compose

If you prefer using Docker:
```bash
# 1. Install Docker & Docker Compose
curl -fsSL https://get.docker.com | sh

# 2. Launch Gola PDF container
docker compose up -d --build
```
Your container will automatically start, monitor health via `/api/health`, and restart if anything fails.

---

## 5. 🛠️ Useful Management & Monitoring Commands

- **Check server status**:
  ```bash
  systemctl status golapdf
  ```
- **View live server logs in real time**:
  ```bash
  journalctl -u golapdf -f
  ```
- **Restart the application**:
  ```bash
  systemctl restart golapdf
  ```
- **Test and reload Nginx**:
  ```bash
  nginx -t && systemctl reload nginx
  ```
- **Renew SSL Certificate (automatically renewed by certbot timer)**:
  ```bash
  certbot renew --dry-run
  ```

---

## 6. 💰 Google AdSense & Search Console Setup

Once your domain is live:
1. **Google Search Console**:
   - Add property: `https://golapdf.com`
   - Submit sitemap: `https://golapdf.com/sitemap.xml`
2. **Google AdSense**:
   - Submit your site for approval.
   - When approved, open `/var/www/golapdf/static/ads.txt` and replace `pub-0000000000000000` with your actual AdSense Publisher ID.
