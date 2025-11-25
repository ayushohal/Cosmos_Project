# 🏗️ Automated Formwork Panel Optimization Tool

## 📝 About

This tool extracts wall dimensions and casting data from PDF plans to calculate the optimal arrangement of formwork panels. It features **parallel panel placement**, **dynamic splitting/merging of IC/EC logic**, and an **inventory management system** to maximize reuse efficiency between primary and secondary castings.

---

## 🛠️ Setup

1.  **Clone the GitHub repository:**
    ```bash
    git clone <repository-url>
    ```

2.  **Install the required dependencies from the `requirements.txt` file:**
    ```bash
    pip install -r requirements.txt
    ```

---

## 🚀 Usage

1.  **Run the application:**
    ```bash
    python app.py
    ```
    The application will provide a link to the local host (e.g., `http://127.0.0.1:5000/`). Open this link in your web browser.

2.  **Upload** the `final_demo.pdf` file.

3.  Select the **"Primary Casting"** option.

4.  Click **"Optimize"**.

5.  Once optimization is complete, generate and download the Excel file by clicking the **"Download Excel"** option.
