# MailMinerGraph

> [!IMPORTANT] 
> This repository, including this README, was entirely AI generated.

MailMinerGraph extracts **email metadata** from a **Microsoft 365 mailbox** using the **Microsoft Graph API**. It collects details like **senders, recipients, dates, folders, and more**, exporting them to a **CSV file** for further analysis.

## 🚀 Installation

1. **Clone the repository**
   ```sh
   git clone https://github.com/YOUR-USERNAME/MailMinerGraph.git
   cd MailMinerGraph
   ```

2. **Create a virtual environment (optional but recommended)**
   ```sh
   python -m venv .venv
   source .venv/bin/activate  # macOS/Linux
   .venv\Scripts\activate   # Windows
   ```

3. **Install dependencies**
   ```sh
   pip install -r requirements.txt
   ```

4. **Ensure OpenSSL Compatibility (macOS Users)**
   If you see a warning about `LibreSSL` when running the script, install OpenSSL to avoid potential SSL/TLS issues:
   ```sh
   brew install openssl
   ```
   Then, ensure your virtual environment uses the correct OpenSSL version:
   ```sh
   export LDFLAGS="-L$(brew --prefix openssl)/lib"
   export CPPFLAGS="-I$(brew --prefix openssl)/include"
   export PKG_CONFIG_PATH="$(brew --prefix openssl)/lib/pkgconfig"
   python -m venv .venv --clear
   source .venv/bin/activate
   pip install -U pip setuptools wheel
   pip install -r requirements.txt
   ```
   To verify the correct OpenSSL version:
   ```sh
   python -c "import ssl; print(ssl.OPENSSL_VERSION)"
   ```
   If the output shows OpenSSL **1.1.1+ or 3.x**, you’re all set!

   Install of installing OpenSSL you could also simply downgrade the version of **urllib**. Probably only do this if you're in a virtual environment.

   ```sh
   pip install "urllib3<2"
   ```

## 🔑 Setup & Authentication

1. **Register an Azure AD App**
   - Go to [Azure Portal](https://portal.azure.com/)
   - Register an **app** under **Azure Active Directory**
   - Add **Microsoft Graph API permissions**:
     - `Mail.Read`
     - `Mail.ReadWrite`
   - Generate a **client secret** and note the **client ID, tenant ID, and secret**

2. **Configure the Script**
   - Copy `example.config.json` to `config.json`
   - Fill in the values from Azure:
     ```json
     {
       "client_id": "YOUR_CLIENT_ID",
       "client_secret": "YOUR_CLIENT_SECRET",
       "tenant_id": "YOUR_TENANT_ID",
       "email": "your@email.com"
     }
     ```

## 📨 Running the Script

To fetch and save email metadata:
```sh
python fetch-emails.py
```
- Output is saved in `email_metadata.csv`
- Errors (if any) are logged in `errors.log`

## 📊 What Can You Do With This Data?

- **Analyze your email habits** (who you email most, how often, etc.)
- **Track spam trends** over time
- **Find old conversations** and categorize your inbox
- **Visualize email volume** over different periods

---

📬 **Start mining your email history today!** 🚀
