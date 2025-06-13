import os
import pandas as pd
from django.shortcuts import render
from django.core.files.storage import FileSystemStorage
import matplotlib.pyplot as plt
import uuid
import requests
from bs4 import BeautifulSoup
from django.conf import settings


def home(request):  
    results = []
    download_url = None

    if request.method == "POST" and 'file' in request.FILES:
        uploaded_file = request.FILES['file']
        file_ext = os.path.splitext(uploaded_file.name)[1].lower()

        if file_ext not in ['.xls', '.xlsx']:
            results.append({"Question": "Error", "Result": "❌ Uploaded file is not a valid Excel (.xls or .xlsx) file."})
            return render(request, "home.html", {
                "results": results,
                "download_url": None
            })

        fs = FileSystemStorage()
        filename = fs.save(uploaded_file.name, uploaded_file)
        file_path = os.path.join(settings.MEDIA_ROOT, filename)


        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_dict = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}

            results = run_all_checks(sheet_dict)

            # Save results
            output_filename = f"QA_Report_{uuid.uuid4().hex}.xlsx"
            output_path = os.path.join(settings.MEDIA_ROOT, output_filename)
            save_results_to_excel(results, output_path)
            download_url = os.path.join(settings.MEDIA_URL, output_filename)

        except Exception as e:
            results.append({"Question": "Error", "Result": str(e)})

    return render(request, "home.html", {
        "results": results,
        "download_url": download_url
    })


def run_analysis(matched_question, df):
    if df is None:
        return "No data to analyze."

    q = matched_question.lower()

    try:
        print(df.head(5))
        # 1. Primary Conversion Action
        if "only one primary conversion action" in q:
            # Normalize column names
            df.columns = df.columns.str.strip()

            if {'Conversion Action Category', 'Conversion Action Primary for Goal'}.issubset(df.columns):
                # Filter for rows where conversion action includes "purchase"
                purchase_rows = df[df['Conversion Action Category'].astype(str).str.contains("purchase", case=False, na=False)]
                
                # Count how many are marked as TRUE for primary conversion goal
                primary_true_count = purchase_rows['Conversion Action Primary for Goal'].astype(str).str.upper().eq("TRUE").sum()

                if primary_true_count == 1:
                    return "Yes – Only one primary conversion action is marked under 'Purchase'."
                elif primary_true_count > 1:
                    return f"No – {primary_true_count} primary conversion actions are marked under 'Purchase'."
                else:
                    return "No – No primary conversion action is marked under 'Purchase'."
            
            return "Required columns 'Conversion Action' and 'Conversion Action Primary for Goal' are missing."


        # 2. Primary Conversion "Purchase" capturing conversions
        elif "purchase" in q and "conversions and revenue" in q:
            # Normalize column names
            df.columns = df.columns.str.strip()

            if {'All Conversions', 'All Conversions Value'}.issubset(df.columns):
                # Convert to string before checking .str.contains()
                df['All Conversions'] = df['All Conversions'].astype(str)

                # Filter rows where the conversion name includes "purchase"
                filtered = df[df['All Conversions'].str.contains("purchase", case=False, na=False)]

                # Check for additional optional 'Conversions' column
                total_conversions = filtered['Conversions'].sum() if 'Conversions' in filtered.columns else 0
                total_value = filtered['All Conversions Value'].fillna(0).sum()

                if not filtered.empty and (total_conversions > 0 or total_value > 0):
                    return "Yes – 'Purchase' is capturing conversions and revenue."
                else:
                    return "No – 'Purchase' is not capturing any conversions or revenue."
            
            return "Required columns 'All Conversions' and 'All Conversions Value' are missing."



        # 3. Campaign names consistent
        elif "campaign names consistent" in q:
            # Normalize column names
            df.columns = df.columns.str.strip()
            
            if 'Campaign Name' in df.columns:
                inconsistent = df[~df['Campaign Name'].fillna("").str.startswith("NX_")]
                
                if inconsistent.empty:
                    return "All campaign names are consistent and start with 'NX_'."
                else:
                    summary = inconsistent[['Campaign Name']].drop_duplicates()
                    html = summary.to_html(index=False, classes="table")
                    return f"{len(summary)} campaign(s) do not start with 'NX_':<br>{html}"
            
            return "'Campaign Name' column is missing."

        # 4. % ad groups with >20 keywords
        elif "ad groups" in q and "20 keywords" in q:
            # Normalize column names
            df.columns = df.columns.str.strip()
            
            if {'Adgroup Name', 'Keyword Name'}.issubset(df.columns):
                group_counts = df.groupby('Adgroup Name')['Keyword Name'].count()
                more_than_20 = (group_counts > 20).sum()
                total = len(group_counts)
                pct = (more_than_20 / total) * 100 if total else 0

                # Pie chart
                fig, ax = plt.subplots()
                ax.pie(
                    [more_than_20, total - more_than_20],
                    labels=[">20 Keywords", "20 or Fewer"],
                    autopct='%1.1f%%',
                    colors=["#ff9999", "#66b3ff"]
                )
                ax.axis('equal')
                chart_id = f"chart_{uuid.uuid4().hex}.png"
                path = f"static/{chart_id}"
                plt.savefig(path, bbox_inches='tight')
                plt.close()

                return f"{pct:.1f}% of ad groups have more than 20 keywords.<br><img src='/{path}' width='400'/>"
            
            return "Required columns 'Adgroup Name' or 'Keyword Name' are missing."


        # 5. Impression Share lost due to budget
        elif "impression share" in q and "budget" in q:
            # Normalize column names
            df.columns = df.columns.str.strip()
            
            required_cols = {'Campaign Name', 'Campaign Type', 'Campaign Status', 'Conversions', 'Search Budget Lost Impression Share'}
            if required_cols.issubset(df.columns):
                filtered = df[
                    (df['Campaign Type'].str.upper().isin(['SEARCH', 'DISPLAY'])) &
                    (df['Conversions'].fillna(0) > 0) &
                    (df['Search Budget Lost Impression Share'].fillna(0) > 10)
                ]
                if filtered.empty:
                    return "No Search or Display campaigns with conversions are losing more than 10% Impression Share due to budget."
                else:
                    summary = filtered[['Campaign', 'Campaign Type', 'Conversions', 'Search Budget Lost Impression Share']].drop_duplicates()
                    html = summary.to_html(index=False, classes="table")
                    return f"{len(summary)} campaigns with conversions are losing over 10% Impression Share due to budget:<br>{html}"
            return "Required columns 'Campaign', 'Campaign Type', 'Conversions', or 'Search Budget Lost Impression Share' are missing."


        # 6. Legacy BMM keywords
        elif "legacy bmm keywords" in q:
            # Normalize column names by stripping whitespace
            df.columns = df.columns.str.strip()
            
            if {'Keyword Name', 'Campaign Name', 'Adgroup Name'}.issubset(df.columns):
                bmm = df[df['Keyword Name'].str.contains(r"\+", na=False)]
                if bmm.empty:
                    return "No legacy BMM keywords found."
                else:
                    summary = bmm[['Keyword Name', 'Campaign Name', 'Adgroup Name']].drop_duplicates()
                    html = summary.to_html(index=False, classes="table")
                    return f"{len(summary)} legacy BMM keywords found:<br>{html}"
            return "Required columns 'Keyword', 'Campaign', or 'Ad group' are missing."


        # 7. Active search ad groups with no conversions (90 days)
        elif "search ad groups" in q and ("no conversions" in q or "not had any conversions" in q):
            if {'Adgroup Type', 'Conversions', 'Campaign Name', 'Adgroup Status'}.issubset(df.columns):
                filtered = df[
                    (df['Adgroup Type'].str.upper() == "SEARCH_STANDARD") &
                    (df['Conversions'].fillna(0) == 0)
                ]
                if filtered.empty:
                    return "All active search ad groups have had at least one conversion in the last 90 days."
                else:
                    summary = filtered[['Campaign Name', 'Adgroup Name', 'Adgroup Status', 'Conversions']].drop_duplicates()
                    html = summary.to_html(index=False, classes="table")
                    return f"{len(summary)} active search ad groups had 0 conversions in the last 90 days:<br>{html}"
            return "Required columns 'Adgroup Type', 'Conversions', 'Campaign', or 'Ad group' are missing."


        # 8. Seasonal keywords
        elif "seasonal keywords" in q:
            if 'Keyword' in df.columns:
                seasonal_terms = ["holiday", "black friday", "back to school", "christmas"]
                seasonal = df[df['Keyword'].str.contains('|'.join(seasonal_terms), case=False, na=False)]
                return f"{len(seasonal)} seasonal keywords found." if not seasonal.empty else "No irrelevant seasonal keywords."
            return "'Keyword' column missing."

        # 9. Low search volume keywords
        elif "low search volume" in q or "rarely_served" in q:
            # Normalize column names
            df.columns = df.columns.str.strip()

            required_cols = {'Status Reason', 'Campaign Name', 'Adgroup Name', 'Keyword Name', 'Keyword MatchType'}
            if required_cols.issubset(df.columns):
                low_volume = df[df['Status Reason'].astype(str).str.upper() == "RARELY_SERVED"]
                
                if low_volume.empty:
                    return "No active keywords are marked as low search volume (RARELY_SERVED)."
                else:
                    summary = low_volume[['Campaign Name', 'Adgroup Name', 'Keyword Name', 'Keyword MatchType']].drop_duplicates()
                    html = summary.to_html(index=False, classes="table")
                    return f"{len(summary)} keyword(s) are low search volume and rarely served:<br>{html}"
            
            return "Required columns are missing: 'Status Reason', 'Campaign', 'Adgroup Name', 'Keyword', or 'Match Type'."


        # 10. Negative dynamic targeting in DSAs
        elif "negative dynamic targeting" in q:
            if 'Campaign type' in df.columns and 'Dynamic ad target' in df.columns:
                dsa = df[df['Campaign type'].str.contains("Dynamic", case=False, na=False)]
                missing = dsa[dsa['Dynamic ad target'].isnull()]
                return "All DSAs have negative targeting set." if missing.empty else f"{len(missing)} DSAs missing targeting rules."
            return "Columns missing."

        # 11. Final URL relevance
        elif ("landing pages" in q and "final url" in q):
            # Normalize column names
            df.columns = df.columns.str.strip()

            required_cols = {'Keyword Final URLs', 'Adgroup Type', 'Campaign Name', 'Keyword Name'}
            if required_cols.issubset(df.columns):
                # Filter: Final URL is missing AND Adgroup Type is NOT DISPLAY_STANDARD
                broken = df[df['Keyword Final URLs'].isnull() & (df['Adgroup Type'].str.upper() != 'DISPLAY_STANDARD')]

                if broken.empty:
                    return "All relevant keywords have Final URLs (landing pages)."
                else:
                    summary = broken[['Campaign Name', 'Adgroup Name', 'Keyword Name', 'Keyword Final URLs', 'Status Reason']].drop_duplicates()
                    html = summary.to_html(index=False, classes="table")
                    return f"{len(summary)} keyword(s) are missing Final URLs (landing pages):<br>{html}"
            
            return "Required columns missing: 'Final URL', 'Adgroup Type', 'Campaign', or 'Keyword'."



        # 12. Broken links / redirections
        elif "broken links" in q or "redirects" in q:
            # Normalize column names
            df.columns = df.columns.str.strip()

            required_cols = {'Keyword Final URLs', 'Adgroup Type', 'Campaign Name', 'Keyword Name'}
            if required_cols.issubset(df.columns):
                filtered = df[(df['Keyword Final URLs'].notna()) & (df['Adgroup Type'].str.upper() != 'DISPLAY_STANDARD')]

                broken_urls = []
                checked_urls = set()

                for url in filtered['Keyword Final URLs'].unique():
                    if url in checked_urls:
                        continue
                    checked_urls.add(url)
                    try:
                        r = requests.head(url, timeout=3, allow_redirects=True)
                        # Sometimes HEAD not allowed, fallback to GET
                        if r.status_code == 405:
                            r = requests.get(url, timeout=3, allow_redirects=True)
                        if r.status_code == 404:
                            broken_urls.append(url)
                    except Exception:
                        broken_urls.append(url)

                if not broken_urls:
                    return "All URLs are working (no 404 errors detected)."
                else:
                    # Get details for broken URLs
                    broken_df = filtered[filtered['Keyword Final URLs'].isin(broken_urls)][['Campaign Name', 'Adgroup Name', 'Keyword Name', 'Keyword Final URLs']].drop_duplicates()
                    html = broken_df.to_html(index=False, classes="table")
                    return f"{len(broken_df)} keywords have final URLs returning 404 error:<br>{html}"

            return "Required columns missing: 'Final URL', 'Adgroup Type', 'Campaign', or 'Keyword'."


        # 13. Legacy ETAs
        elif "legacy expanded text ads" in q or "legacy" in q and "account" in q:
            df.columns = df.columns.str.strip()

            required_cols = {'Ad Type', 'Campaign Name', 'Adgroup Name'}
            if required_cols.issubset(df.columns):
                # Filter for ETA or Text Ad (case-insensitive)
                eta = df[(df['Ad Type'].str.upper() == 'EXPANDED_DYNAMIC_SEARCH_AD')]

                if eta.empty:
                    return "No legacy ETAs found."
                else:
                    # Group by Campaign and Adgroup, count ETAs
                    grouped = eta.groupby(['Campaign Name', 'Adgroup Name']).size().reset_index(name='ETA Count')
                    html = grouped.to_html(index=False, classes="table")
                    total_count = len(eta)
                    return f"{total_count} legacy ETAs found.<br>{html}"

            return "Required columns missing: 'Ad type', 'Campaign', or 'Adgroup Name'."


        # 14. At least one RSA per ad group with excellent ad strength
        elif "rsa per ad group" in q or "ad strength" in q and "excellent" in q:
            df.columns = df.columns.str.strip()

            required_cols = {'Adgroup Name', 'Ad Type', 'Ad Strength', 'Campaign Name'}
            if required_cols.issubset(df.columns):
                # Filter RSAs
                rsa_df = df[(df['Ad Type'].str.upper() == 'RESPONSIVE_SEARCH_AD')]

                # Group by Ad group and Campaign
                grouped = rsa_df.groupby(['Campaign Name', 'Adgroup Name'])

                # Count total RSAs per ad group
                total_rsa_counts = grouped.size().rename('Total RSA')

                # Count RSAs with Ad Strength 'Excellent'
                excellent_rsa_counts = grouped.apply(lambda x: (x['Ad Strength'].str.upper() == "EXCELLENT").sum()).rename('Excellent RSA')

                # Combine counts into one dataframe
                summary = pd.concat([total_rsa_counts, excellent_rsa_counts], axis=1).reset_index()

                # Count ad groups missing Excellent RSAs
                missing_count = (summary['Excellent RSA'] == 0).sum()

                if missing_count == 0:
                    return f"All ad groups have at least one RSA with excellent ad strength.<br>" + summary.to_html(index=False, classes="table")
                else:
                    missing_summary = summary[summary['Excellent RSA'] == 0]
                    return (f"{missing_count} ad group(s) missing RSAs with excellent ad strength.<br>"
                            f"Summary:<br>{summary.to_html(index=False, classes='table')}<br><br>"
                            f"Ad groups missing excellent RSAs:<br>{missing_summary.to_html(index=False, classes='table')}")
            return "Required columns missing: 'Ad group', 'Ad type', 'Ad Strength', or 'Campaign'."
            


        # 15. RSAs use all headlines/descriptions
        elif "rsa" in q and ("headlines" in q or "descriptions" in q):
            df.columns = df.columns.str.strip()

            required_cols = {'Ad Type', 'RSA Headlines Count', 'RSA Descriptions Count', 'Campaign Name', 'Adgroup Name'}
            if required_cols.issubset(df.columns):
                # Filter Responsive Search Ads
                rsa_df = df[df['Ad Type'].str.contains("Responsive", case=False, na=False)]

                if rsa_df.empty:
                    return "No RSAs found."

                # Convert Headlines and Descriptions to numeric if needed
                rsa_df['RSA Headlines Count'] = pd.to_numeric(rsa_df['RSA Headlines Count'], errors='coerce').fillna(0)
                rsa_df['RSA Descriptions Count'] = pd.to_numeric(rsa_df['RSA Descriptions Count'], errors='coerce').fillna(0)

                # Determine which RSAs meet the full usage criteria
                rsa_df['Pass Criteria'] = (rsa_df['RSA Headlines Count'] >= 15) & (rsa_df['RSA Descriptions Count'] >= 4)

                # Group by Campaign and Adgroup
                grouped = rsa_df.groupby(['Campaign Name', 'Adgroup Name'])

                # Count total RSAs and how many pass
                summary = grouped.agg(
                    Total_RSAs = ('Ad Type', 'count'),
                    Pass_Criteria = ('Pass Criteria', 'sum')
                ).reset_index()

                underused_count = summary[summary['Pass_Criteria'] < summary['Total_RSAs']].shape[0]

                if underused_count == 0:
                    return "All RSAs are using all headline and description slots.<br>" + summary.to_html(index=False, classes="table")
                else:
                    return (f"{underused_count} campaign/adgroup(s) have RSAs underutilizing headlines/descriptions.<br>"
                            + summary.to_html(index=False, classes="table"))
            return "Required columns missing: 'Ad type', 'Headlines', 'Descriptions', 'Campaign', or 'Adgroup Name'."

#do this
        # 16. Ad extensions (sitelinks, callouts, etc.)
        elif "ad extensions" in q:
            df.columns = df.columns.str.strip()

            # List of expected ad extension columns
            extension_cols = ['Campaign Name', 'Campaign Type', 'Feed Item Status', 'Extension Type']

            # Check which of them are actually in the dataframe
            available_cols = [col for col in extension_cols if col in df.columns]

            if not available_cols:
                return "None of the required ad extension columns (Sitelinks, Callouts, Calls, Structured Snippets, Promotions) are present in the data."

            # Count missing/non-empty for each extension
            summary = {}
            sitelinks_count=0
            callouts_count=0
            snippets_count=0
            promos_count=0
            for col in available_cols:
                total = len(df)
                present = df[col].notnull().sum()
                if present == 'sitelinks':
                    sitelinks_count+=1
                elif present=='callouts':
                    callouts_count+=1
                elif present=='snippets':
                    snippets_count+=1
                elif present=='promos':
                    promos_count+=1
                summary[col] = {
                    'Total Rows': total,
                    'With Extension': present,
                    'Missing': total - present,
                    'Coverage %': round((present / total) * 100, 1) if total > 0 else 0.0
                }

            # Convert to DataFrame
            result_df = pd.DataFrame(summary).T.reset_index().rename(columns={'index': 'Extension Type'})

            return (
                "Ad Extension Implementation Summary:<br>" +
                result_df.to_html(index=False, classes="table")
            )

#do
        # 17. Sitelinks have descriptions
        elif "sitelink descriptions" in q:
            if 'Sitelink description' in df.columns:
                missing = df['Sitelink description'].isnull().sum()
                return "All sitelinks have descriptions." if missing == 0 else f"{missing} sitelinks missing descriptions."
            return "Column missing."
        
#do
        # 18. Affinity/In-Market audience in Observation
        elif "affinity" in q or "in-market" in q:
            if 'Audience setting' in df.columns:
                missing = df[~df['Audience setting'].str.contains("Observation", na=False, case=False)]
                return "Observation mode set for all." if missing.empty else f"{len(missing)} entries missing Observation mode."
            return "'Audience setting' column missing."
        
#do
        # 19. Performance Max: audience signals
        elif "performance max" in q and "audience signal" in q:
            if 'Audience signal' in df.columns:
                empty = df['Audience signal'].isnull().sum()
                return "All Performance Max campaigns have audience signals." if empty == 0 else f"{empty} missing audience signals."
            return "Column missing."

#do
        # 20. Performance Max video assets
        elif "performance max" in q and "video" in q:
            if 'Video Asset' in df.columns:
                missing = df['Video Asset'].isnull().sum()
                return "All asset groups have videos." if missing == 0 else f"{missing} asset groups missing videos."
            return "Column missing."
        
        
        # 21. Active display ad groups with no conversions or view-through conversions
        elif "display ad groups" in q and ("no conversions" in q or "view-through conversions" in q):
            df.columns = df.columns.str.strip()

            required_cols = {'Adgroup Type', 'Conversions', 'View Through Conversions', 'Campaign Name', 'Adgroup Name'}
            if required_cols.issubset(df.columns):
                filtered = df[
                    (df['Adgroup Type'].str.upper() == "DISPLAY_STANDARD") &
                    ((df['Conversions'].fillna(0) == 0) | (df['View Through Conversions'].fillna(0) == 0))
                ]

                if filtered.empty:
                    return "All active display ad groups had either conversions or view-through conversions in the last 90 days."
                else:
                    output = filtered[['Campaign Name', 'Adgroup Name', 'Conversions', 'View Through Conversions']].drop_duplicates()
                    return (
                        f"{len(output)} active display ad groups had 0 conversions or view-through conversions:<br>"
                        f"{output.to_html(index=False, classes='table')}"
                    )
            else:
                return "Required columns missing: 'Adgroup Type', 'Conversions', 'View-through Conversions', 'Campaign Name', or 'Adgroup Name'."

        else:
            return "Matched your question but logic for it isn't implemented yet."
        
    except Exception as e:
        return f"Error: {e}"


def load_predefined_questions():
    with open(r"C:\Users\cscpr\Desktop\Internship\questions.txt", 'r') as file:
        questions = [line.strip() for line in file.readlines() if line.strip()]
    return questions


QUESTION_TO_SHEET_MAP = {
    "Is there only one primary conversion action?": "Conversions Tracking Data",
    'If the primary conversion action is "Purchase," is it capturing conversions and revenue properly?': "Conversions Tracking Data",
    "Are campaign names consistent across the account?": "Campaign Data",
    "What percentage of ad groups have more than 20 keywords?": "Keyword Data",
    "Are Search Campaigns or Display Campaigns with conversions losing Impression Share due to budget limitations?": "Campaign Data",
    "Are there any legacy BMM keywords?": "Keyword Data",
    "Are there active search ad groups that have not had any conversions in the last 90 days?": "AdGroup Data",
    "Are there any seasonal keywords, like back-to-school or holiday keywords running that are not relevant to the current season?": "Keyword Data",
    "Are there active keywords with low search volumes that are not receiving enough impressions?": "Keyword Data",
    "Are negative dynamic targeting options set for all Dynamic Search Ads campaigns?": "DSA",
    "Are there landing pages (Final URL) at the keyword level, and are they relevant for the ad message, keywords, and targeting?": "Keyword Data",
    "Are there any broken links or redirections in final URLs?": "Keyword Data",
    "Are there still legacy Expanded Text Ads (ETAs) live in the account?": "Ad Data",
    "Is there at least one RSA per ad group with an ad strength of excellent?": "Ad Data",
    "Are the RSAs leveraging all available headlines (15) and description lines (4)?": "RSA Ad Data",
    "Does the account have ad extensions implemented, such as sitelinks, callouts, calls, structured snippets, and promos?": "Extensions Data",
    "Do all sitelinks have expanded sitelink text filled in (descriptions)?": "Extensions",
    'Are Affinity and In-Market audiences applied to the campaigns in "Observation" mode at Campaign Level?': "Audiences",
    "Have both customer data and interests been included in the audience signal for Performance Max?": "Campaigns",
    "Do Performance Max campaign asset groups have at least one customized video?": "Campaigns",
    "Are there active display ad groups with no conversions or view-through conversions in the last 90 days?": "AdGroup Data"
}


def run_all_checks(sheet_dict):
    all_questions = load_predefined_questions()
    results = []

    for question in all_questions:
        sheet_name = QUESTION_TO_SHEET_MAP.get(question)

        if not sheet_name:
            results.append({
                "Question": question,
                "Result": "❌ No sheet mapping defined for this question."
            })
            continue

        df = sheet_dict.get(sheet_name)
        if df is None:
            results.append({
                "Question": question,
                "Result": f"❌ Sheet '{sheet_name}' not found in uploaded Excel file."
            })
            continue

        try:
            result = run_analysis(question, df)
            results.append({
                "Question": question,
                "Result": f"✅ {sheet_name}: {result}"
            })
        except Exception as e:
            results.append({
                "Question": question,
                "Result": f"❌ Error analyzing '{sheet_name}': {str(e)}"
            })

    return results




def strip_html(raw_html):
    soup = BeautifulSoup(raw_html, "html.parser")
    return soup.get_text(separator="\n")

def save_results_to_excel(results, output_path):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary = []

        for i, item in enumerate(results):
            question = item["Question"]
            result = item["Result"]
            sheet_name = f"Q{i+1:02d}"[:31]

            if "<table" in result:
                try:
                    tables = pd.read_html(result)
                    if tables:
                        tables[0].to_excel(writer, sheet_name=sheet_name, index=False)
                        summary.append({
                            "Question": question,
                            "Result Summary": f"✅ See sheet '{sheet_name}'"
                        })
                except Exception as e:
                    summary.append({
                        "Question": question,
                        "Result Summary": f"⚠ Error rendering table: {e}"
                    })
            else:
                # Clean simple text
                clean = strip_html(result)
                summary.append({
                    "Question": question,
                    "Result Summary": clean
                })

        pd.DataFrame(summary).to_excel(writer, sheet_name="Summary", index=False)



