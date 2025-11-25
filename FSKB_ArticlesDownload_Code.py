
# -*- coding: utf-8 -*-
"""
Enhanced version of the user's existing `downloadFSKB.py` that preserves all original
functionality and console/reporting, while ADDING the following capabilities:

1) Full article content download (HTML, PDF, and TXT) per article.
2) Per-article directory layout: knowledge_base/<id>_<title>/ with attachments inside.
3) Rich per-article metadata JSON.
4) Summary of articles and attachments downloaded.

Notes:
- The original code is kept intact; new methods and an additional call in `main()` are
  appended to achieve the recommended features without removing the original behavior.
- PDF generation uses WeasyPrint. Ensure WeasyPrint and its system dependencies are
  installed on the runtime environment. If WeasyPrint is unavailable, the code will
  skip PDF generation gracefully and log an error.
"""

import base64
import requests
import json
import csv
import os
import logging
import pandas as pd
from datetime import datetime
from urllib.parse import urljoin
from tabulate import tabulate
import urllib3

# Additional imports for enhanced capabilities
import re
import html as html_lib  # to avoid name clash with any 'html' variable

# Try to import WeasyPrint; handle absence gracefully inside methods
try:
    from weasyprint import HTML, CSS
    from weasyprint.text.fonts import FontConfiguration
    WEASYPRINT_AVAILABLE = True
except Exception:
    WEASYPRINT_AVAILABLE = False

# Suppress SSL warnings if verify=False (preserving original behavior)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class FreshServiceKBExtractor:
    def __init__(self, domain, api_key):
        self.domain = domain
        self.api_key = api_key
        self.base_url = f"https://{domain}.freshservice.com/api/v2"
        auth_string = f"{api_key}:X"
        encoded_auth = base64.b64encode(auth_string.encode()).decode()
        self.headers = {
            'Authorization': f'Basic {encoded_auth}',
            'Content-Type': 'application/json'
        }
        self.setup_logging()

    def setup_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('kb_extraction.log'),
                logging.StreamHandler()
            ]
        )

    def validate_connection(self) -> bool:
        url = f"{self.base_url}/agents/me"
        try:
            response = requests.get(url, headers=self.headers, timeout=10, verify=False)
            if response.status_code == 200:
                print("Connection successful!")
                return True
            elif response.status_code == 403:
                print("Permission denied: Check if your API key user has KB access.")
                return False
            else:
                print(f"Connection failed: {response.status_code} - {response.text}")
                return False
        except requests.RequestException as e:
            print(f"Error: {e}")
            return False

    def get_categories(self):
        url = f"{self.base_url}/solutions/categories"
        response = requests.get(url, headers=self.headers, verify=False)
        if response.status_code == 200:
            return response.json().get("categories", [])
        elif response.status_code == 403:
            logging.error("Permission denied for categories. Check KB access.")
            return []

    def get_folders(self, category_id):
        url = f"{self.base_url}/solutions/folders"
        params = {"category_id": category_id}
        response = requests.get(url, headers=self.headers, params=params, verify=False)
        if response.status_code == 200:
            return response.json().get("folders", [])
        elif response.status_code == 403:
            logging.error(f"Permission denied for folders in category {category_id}.")
            return []

    def get_articles(self, folder_id=None, category_id=None):
        articles = []
        page = 1
        while True:
            url = f"{self.base_url}/solutions/articles"
            params = {'page': page, 'per_page': 100}
            if folder_id:
                params["folder_id"] = folder_id
            elif category_id:
                params["category_id"] = category_id
            response = requests.get(url, headers=self.headers, params=params, verify=False)
            if response.status_code == 200:
                batch = response.json().get("articles", [])
                articles.extend(batch)
                if len(batch) < 100:
                    break
                page += 1
            elif response.status_code == 403:
                logging.error("Permission denied for articles. Check KB access.")
                break
            else:
                logging.error(f"Error fetching articles: {response.status_code}")
                break
        return articles

    def get_article_details(self, article_id):
        url = f"{self.base_url}/solutions/articles/{article_id}"
        try:
            response = requests.get(url, headers=self.headers, verify=False)
            if response.status_code == 200:
                return response.json().get('article')
            else:
                logging.error(f"Error fetching article details for {article_id}: {response.status_code}")
                return None
        except requests.RequestException as e:
            logging.error(f"Error fetching article details for {article_id}: {e}")
            return None

    def display_table(self, data, headers, title):
        if data:
            print(f"\n{title}")
            print(tabulate(data, headers=headers, tablefmt="simple"))
        else:
            print(f"\n{title} - No data found.")

    def sanitize_filename(self, filename):
        invalid_chars = '<>:"/\\\n?*'
        filename = filename.strip()  # Remove leading/trailing spaces
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename[:250]

    def download_attachment(self, attachment, download_path):
        try:
            download_url = attachment.get('attachment_url')
            if not download_url:
                logging.warning(f"No download URL for attachment {attachment.get('id')}")
                return False
            if not download_url.startswith('http'):
                download_url = urljoin(self.base_url, download_url)
            file_headers = self.headers.copy()
            file_headers.pop('Content-Type', None)
            response = requests.get(download_url, headers=file_headers, stream=True, verify=False)
            response.raise_for_status()
            os.makedirs(os.path.dirname(download_path), exist_ok=True)
            with open(download_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            logging.info(f"Downloaded attachment: {download_path}")
            return True
        except requests.exceptions.RequestException as e:
            logging.error(f"Error downloading attachment {attachment.get('id')}: {e}")
            return False

    def download_all_attachments(self, articles, base_download_dir="attachments"):
        total_attachments = 0
        successful_downloads = 0
        for article in articles:
            article_id = article['id']
            article_title = self.sanitize_filename(article.get('title', f'article_{article_id}'))
            detailed_article = self.get_article_details(article_id)
            if not detailed_article:
                continue
            attachments = detailed_article.get('attachments', [])
            if attachments:
                logging.info(f"Found {len(attachments)} attachments for article: {article_title}")
                article_dir = os.path.normpath(os.path.join(base_download_dir, f"{article_id}_{article_title}"))
                os.makedirs(article_dir, exist_ok=True)
                for attachment in attachments:
                    attachment_id = attachment.get('id')
                    attachment_name = self.sanitize_filename(attachment.get('name', 'unknown'))
                    filename = f"{attachment_id}_{attachment_name}"
                    download_path = os.path.join(article_dir, filename)
                    total_attachments += 1
                    if self.download_attachment(attachment, download_path):
                        successful_downloads += 1
        logging.info(f"Attachment download summary: {successful_downloads}/{total_attachments} successful")

    def export_to_json(self, articles, filename=None):
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"freshservice_kb_export_{timestamp}.json"
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(articles, f, indent=2, ensure_ascii=False)
        logging.info(f"Exported {len(articles)} articles to {filename}")
        return filename

    def export_to_csv(self, articles, filename=None):
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"freshservice_kb_export_{timestamp}.csv"
        if articles:
            fieldnames = set()
            for article in articles:
                fieldnames.update(article.keys())
            fieldnames = sorted(list(fieldnames))
            with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(articles)
            logging.info(f"Exported {len(articles)} articles to {filename}")
        return filename

    def export_to_excel(self, articles, filename=None):
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"freshservice_kb_export_{timestamp}.xlsx"
        df = pd.DataFrame(articles)
        # Using default engine; ensure 'openpyxl' is installed in the environment.
        df.to_excel(filename, index=False)
        logging.info(f"Exported {len(articles)} articles to {filename}")
        return filename

    def generate_summary_report(self, articles):
        total_articles = len(articles)
        published = len([a for a in articles if a.get('status') == 1])
        draft = len([a for a in articles if a.get('status') == 2])
        report = {
            "timestamp": datetime.now().isoformat(),
            "summary": {
                "total_articles": total_articles,
                "published_articles": published,
                "draft_articles": draft
            }
        }
        return report

    # Added methods: content & PDF

    def create_article_folder(self, article, base_directory="knowledge_base"):
        """Create a per-article folder and return the path."""
        article_id = article['id']
        article_title = self.sanitize_filename(article.get('title', f'article_{article_id}'))
        folder_name = f"{article_id}_{article_title}"
        folder_path = os.path.join(base_directory, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        os.makedirs(os.path.join(folder_path, 'attachments'), exist_ok=True)
        return folder_path

    def html_to_text(self, html_content):
        """Convert HTML content to plain text with better formatting."""
        if not html_content:
            return ""
        replacements = [
            (r'<br\s*/?>', '\n'),
            (r'<p>', '\n'),
            (r'</p>', '\n\n'),
            (r'<h[1-6]>', '\n\n** '),
            (r'</h[1-6]>', ' **\n\n'),
            (r'<li>', '• '),
            (r'</li>', '\n'),
            (r'<ul>', '\n'),
            (r'</ul>', '\n'),
            (r'<ol>', '\n'),
            (r'</ol>', '\n'),
            (r'<strong>', '**'),
            (r'</strong>', '**'),
            (r'<em>', '*'),
            (r'</em>', '*'),
            (r'<code>', '`'),
            (r'</code>', '`'),
            (r'<pre>', '\n```\n'),
            (r'</pre>', '\n```\n'),
            (r'<blockquote>', '\n> '),
            (r'</blockquote>', '\n'),
        ]
        text = html_content
        for pattern, replacement in replacements:
            text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
        text = re.sub(r'<[^>]+>', '', text)  # remove other tags
        text = html_lib.unescape(text)       # decode entities
        text = re.sub(r'\n\s*\n', '\n\n', text)
        text = re.sub(r'[ \t]+', ' ', text)
        return text.strip()

    def create_html_document(self, article, content, for_pdf=False):
        """Create styled HTML for saving or PDF conversion."""
        attachments = article.get('attachments', [])
        attachments_html = ""
        if attachments and for_pdf:
            items = ''.join([f'<div class="attachment-item">• {html_lib.escape(att.get("name", "Unnamed"))}</div>' for att in attachments])
            attachments_html = f'''
            <div class="attachments-section">
              <div class="attachments-title">📎 Article Attachments ({len(attachments)}):</div>
              {items}
            </div>
            '''
        status_map = {1: "Published", 2: "Draft"}
        status = status_map.get(article.get('status'), "Unknown")
        html_template = f"""
        <!DOCTYPE html>
        <html lang="en">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>{html_lib.escape(article.get('title', 'Untitled'))}</title>
          <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 24px; }}
            .header {{ border-bottom: 3px solid #007cba; padding-bottom: 12px; margin-bottom: 24px; }}
            h1 {{ color: #007cba; font-size: 24pt; margin-bottom: 10px; }}
            .metadata {{ background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 5px; padding: 12px; margin-bottom: 20px; font-size: 10pt; }}
            .metadata-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }}
            .metadata-label {{ font-weight: bold; color: #555; }}
            .content {{ margin-top: 10px; font-size: 11pt; }}
            .content h2 {{ color: #007cba; font-size: 14pt; margin-top: 22px; margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 5px; }}
            .content h3 {{ color: #555; font-size: 12pt; margin-top: 18px; margin-bottom: 8px; }}
            .content p {{ margin-bottom: 12px; text-align: justify; }}
            .content ul, .content ol {{ margin-bottom: 12px; margin-left: 20px; }}
            .content table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 9pt; }}
            .content th, .content td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            .content th {{ background-color: #f8f9fa; font-weight: bold; }}
            .content img {{ max-width: 100%; height: auto; margin: 10px 0; }}
            .attachments-section {{ margin-top: 20px; padding: 12px; background-color: #f8f9fa; border-radius: 5px; border-left: 4px solid #007cba; }}
            .attachments-title {{ font-weight: bold; margin-bottom: 8px; color: #007cba; }}
            .footer {{ margin-top: 24px; padding-top: 12px; border-top: 1px solid #eee; font-size: 9pt; color: #666; text-align: center; }}
            code {{ background-color: #f4f4f4; padding: 2px 4px; border-radius: 3px; font-family: 'Courier New', monospace; font-size: 9pt; }}
            pre {{ background-color: #f8f9fa; padding: 10px; border-radius: 5px; overflow-x: auto; font-family: 'Courier New', monospace; font-size: 9pt; border: 1px solid #e9ecef; }}
          </style>
        </head>
        <body>
          <div class="header"><h1>{html_lib.escape(article.get('title', 'Untitled'))}</h1></div>
          <div class="metadata">
            <div class="metadata-grid">
              <div><span class="metadata-label">Article ID:</span> {article.get('id', 'N/A')}</div>
              <div><span class="metadata-label">Status:</span> {status}</div>
              <div><span class="metadata-label">Created:</span> {article.get('created_at', 'N/A')}</div>
              <div><span class="metadata-label">Updated:</span> {article.get('updated_at', 'N/A')}</div>
              <div><span class="metadata-label">Views:</span> {article.get('view_count', 0)}</div>
              <div><span class="metadata-label">Helpful Votes:</span> {article.get('thumbs_up', 0)}</div>
              <div><span class="metadata-label">Not Helpful:</span> {article.get('thumbs_down', 0)}</div>
              <div><span class="metadata-label">Attachments:</span> {len(attachments)}</div>
            </div>
          </div>
          <div class="content">{content}</div>
          {attachments_html}
          <div class="footer">
            <p>Knowledge Base Article • {self.domain}.freshservice.com</p>
            <p>Downloaded on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} • Article URL: {article.get('url', 'N/A')}</p>
          </div>
        </body>
        </html>
        """
        return html_template

    def create_pdf_document(self, article, content, output_path):
        """Create a PDF document from article content using WeasyPrint."""
        if not WEASYPRINT_AVAILABLE:
            logging.error("WeasyPrint not available; skipping PDF generation.")
            return False
        try:
            html_content = self.create_html_document(article, content, for_pdf=True)
            font_config = FontConfiguration()
            pdf_css = CSS(string=f'''
                @page {{
                    size: A4; margin: 1in;
                    @top-left {{ content: "{article.get('title', 'Knowledge Base Article')}"; font-size: 10pt; color: #666; }}
                    @top-right {{ content: "Page " counter(page) " of " counter(pages); font-size: 10pt; color: #666; }}
                    @bottom-left {{ content: "Downloaded from FreshService • {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"; font-size: 8pt; color: #999; }}
                }}
            ''', font_config=font_config)
            HTML(string=html_content).write_pdf(output_path, stylesheets=[pdf_css], font_config=font_config)
            logging.info(f"PDF generated successfully: {output_path}")
            return True
        except Exception as e:
            logging.error(f"Error generating PDF for article {article.get('id')}: {e}")
            return False

    def extract_article_metadata(self, article):
        """Extract relevant metadata from detailed article object."""
        return {
            "id": article.get("id"),
            "title": article.get("title"),
            "status": "Published" if article.get("status") == 1 else "Draft",
            "folder_id": article.get("folder_id"),
            "category_id": article.get("category_id"),
            "created_at": article.get("created_at"),
            "updated_at": article.get("updated_at"),
            "view_count": article.get("view_count"),
            "thumbs_up": article.get("thumbs_up"),
            "thumbs_down": article.get("thumbs_down"),
            "url": article.get("url"),
            "tags": article.get("tags", []),
            "attachments_count": len(article.get("attachments", [])),
            "download_timestamp": datetime.now().isoformat()
        }

    def download_article_content(self, article, folder_path, save_pdf=True, save_html=True, save_text=True):
        """Download and save article content in multiple formats, plus metadata JSON.
        Returns (success: bool, detailed_article: dict or None)
        """
        try:
            article_id = article['id']
            detailed_article = self.get_article_details(article_id)
            if not detailed_article:
                logging.error(f"Could not get detailed content for article {article_id}")
                return False, None
            html_content = detailed_article.get('description', '')
            formats_saved = []

            # Save PDF
            if save_pdf and html_content:
                pdf_file = os.path.join(folder_path, f"{article_id}_article.pdf")
                if self.create_pdf_document(detailed_article, html_content, pdf_file):
                    formats_saved.append("PDF")
                    logging.info(f"Saved PDF article: {pdf_file}")

            # Save HTML
            if save_html and html_content:
                html_file = os.path.join(folder_path, f"{article_id}_article.html")
                with open(html_file, 'w', encoding='utf-8') as f:
                    f.write(self.create_html_document(detailed_article, html_content))
                formats_saved.append("HTML")
                logging.info(f"Saved HTML article: {html_file}")

            # Save Text
            if save_text:
                text_content = self.html_to_text(html_content)
                text_file = os.path.join(folder_path, f"{article_id}_article.txt")
                with open(text_file, 'w', encoding='utf-8') as f:
                    f.write(text_content)
                formats_saved.append("TEXT")
                logging.info(f"Saved text article: {text_file}")

            # Save Metadata JSON
            metadata_file = os.path.join(folder_path, f"{article_id}_metadata.json")
            with open(metadata_file, 'w', encoding='utf-8') as f:
                json.dump(self.extract_article_metadata(detailed_article), f, indent=2, ensure_ascii=False)
            formats_saved.append("JSON")
            logging.info(f"Saved article metadata: {metadata_file}")

            logging.info(f"Successfully saved article in formats: {', '.join(formats_saved)}")
            return True, detailed_article
        except Exception as e:
            logging.error(f"Error downloading article content for {article.get('id')}: {e}")
            return False, None

    def download_articles_and_attachments(self, articles, base_download_dir="knowledge_base",
                                          download_attachments=True, save_pdf=True,
                                          save_html=True, save_text=True):
        """Download article content and attachments using recommended directory layout."""
        total_articles = len(articles)
        successful_articles = 0
        total_attachments = 0
        successful_attachments = 0

        logging.info(f"Starting download of {total_articles} articles to '{base_download_dir}'")

        for index, article in enumerate(articles, 1):
            article_id = article['id']
            article_title = article.get('title', f'article_{article_id}')
            logging.info(f"Processing article {index}/{total_articles}: {article_title}")

            try:
                # Create per-article folder (recommended layout)
                folder_path = self.create_article_folder(article, base_download_dir)

                # Download article content
                ok, detailed_article = self.download_article_content(
                    article, folder_path, save_pdf=save_pdf, save_html=save_html, save_text=save_text
                )
                if ok:
                    successful_articles += 1
                    logging.info(f"Successfully downloaded article content to {folder_path}")
                else:
                    logging.error(f"Failed to download article content for {article_title}")
                    continue

                # Download attachments under the article folder
                if download_attachments and detailed_article:
                    attachments = detailed_article.get('attachments', [])
                    if attachments:
                        logging.info(f"Found {len(attachments)} attachments for article {article_title}")
                        attachment_success_count = 0
                        for att in attachments:
                            att_id = att.get('id')
                            att_name = self.sanitize_filename(att.get('name', 'unknown'))
                            filename = f"{att_id}_{att_name}"
                            download_path = os.path.join(folder_path, 'attachments', filename)
                            total_attachments += 1
                            if self.download_attachment(att, download_path):
                                successful_attachments += 1
                                attachment_success_count += 1
                        logging.info(f"Downloaded {attachment_success_count}/{len(attachments)} attachments for article {article_title}")

                    # annotate article dict (non-destructive)
                    article['has_attachments'] = len(attachments) > 0
                    article['attachment_count'] = len(attachments)
                    article['attachments_downloaded'] = attachment_success_count if attachments else 0
                    article['local_folder'] = folder_path

            except Exception as e:
                logging.error(f"Error processing article {article_title}: {e}")
                continue

        # Summary
        logging.info(f"\n{'='*60}")
        logging.info("DOWNLOAD SUMMARY (enhanced)")
        logging.info(f"{'='*60}")
        logging.info(f"Total articles processed: {total_articles}")
        logging.info(f"Successful article downloads: {successful_articles}")
        logging.info(f"Total attachments found: {total_attachments}")
        logging.info(f"Successful attachment downloads: {successful_attachments}")
        logging.info(f"Download location: {base_download_dir}")

        return {
            'total_articles': total_articles,
            'successful_articles': successful_articles,
            'total_attachments': total_attachments,
            'successful_attachments': successful_attachments,
            'download_directory': base_download_dir
        }


def main():
    DOMAIN = "avh"  # Replace with your Freshservice domain
    API_KEY = "MBMJmETmtwpCYwHwWhAK"  # Replace with your API key
    ATTACHMENTS_DIR = "kb_attachments"

    extractor = FreshServiceKBExtractor(DOMAIN, API_KEY)

    print("Testing Freshservice API connection...")
    if not extractor.validate_connection():
        return

    print("Starting Knowledge Base extraction...")
    categories = extractor.get_categories()

    extractor.display_table(
        [[c["id"], c["name"], c.get("description", ""), c.get("created_at", "")] for c in categories],
        ["ID", "Name", "Description", "Created At"], "Solution Categories"
    )

    all_articles = []
    for category in categories:
        cat_id = category["id"]
        folders = extractor.get_folders(cat_id)
        extractor.display_table(
            [[f["id"], f["name"], f.get("description", ""), f.get("created_at", "")] for f in folders],
            ["ID", "Name", "Description", "Created At"], f"Folders in Category: {category['name']}"
        )
        if folders:
            for folder in folders:
                folder_id = folder["id"]
                articles = extractor.get_articles(folder_id=folder_id)
                extractor.display_table(
                    [[a["id"], a["title"], a.get("status", ""), a.get("created_at", "")] for a in articles],
                    ["ID", "Title", "Status", "Created At"], f"Articles in Folder: {folder['name']}"
                )
                all_articles.extend(articles)
        else:
            articles = extractor.get_articles(category_id=cat_id)
            extractor.display_table(
                [[a["id"], a["title"], a.get("status", ""), a.get("created_at", "")] for a in articles],
                ["ID", "Title", "Status", "Created At"], f"Articles in Category: {category['name']}"
            )
            all_articles.extend(articles)

    if all_articles:
        print(f"\nFound {len(all_articles)} articles in total.")

        json_file = extractor.export_to_json(all_articles)
        csv_file = extractor.export_to_csv(all_articles)
        excel_file = extractor.export_to_excel(all_articles)
        summary = extractor.generate_summary_report(all_articles)
        print(f"\nExported files:\n - {json_file}\n - {csv_file}\n - {excel_file}")
        print("\nSummary Report:")
        print(json.dumps(summary, indent=2))

        print("\nStarting enhanced download: articles (PDF/HTML/TEXT) + per-article attachments...")
        enhanced_result = extractor.download_articles_and_attachments(
            all_articles,
            base_download_dir="knowledge_base",      # recommended layout
            download_attachments=True,
            save_pdf=True,
            save_html=True,
            save_text=True
        )
        print(f"\n{'='*60}")
        print("ENHANCED DOWNLOAD COMPLETED")
        print(f"{'='*60}")
        print(f"Articles successfully downloaded: {enhanced_result['successful_articles']}/{enhanced_result['total_articles']}")
        print(f"Attachments successfully downloaded: {enhanced_result['successful_attachments']}/{enhanced_result['total_attachments']}")
        print(f"All enhanced content saved to: {enhanced_result['download_directory']}")

    else:
        print("No articles found or error occurred.")


if __name__ == "__main__":
    main()
