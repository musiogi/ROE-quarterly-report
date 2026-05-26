import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import os
import base64
import win32com.client as win32

# =====================
# 1. LOAD DATA
# =====================
file_path = "[...]ECB Data Portal long_ROE.xlsx" # TO DO: Add the directory of your file here, after downloading it from the repository
df = pd.read_excel(file_path, sheet_name=1, engine="openpyxl")

# =====================
# 2. CLEAN DATA
# =====================
df["DATE"] = pd.to_datetime(df["DATE"])

df = df[df["DATE"] >= "2020-01-01"]

df["QUARTER"] = df["DATE"].dt.to_period("Q")

df = df.sort_values(["REFERENCE AREA", "DATE"])

# =====================
# 3. SETUP OUTLOOK
# =====================
outlook = win32.Dispatch("Outlook.Application")

# =====================
# 4. SPLIT DATA
# =====================
ssm = df[df["REFERENCE AREA"] == "B01"].sort_values("DATE")
countries = df["REFERENCE AREA"].unique()

latest_ssm_quarter = ssm["QUARTER"].max()

print("Reference quarter:", latest_ssm_quarter)

# =====================
# 5. LOOP COUNTRIES
# =====================
for country in countries:

    if country == "B01":
        continue

    data = df[df["REFERENCE AREA"] == country].sort_values("DATE")

    # =====================
    # CHECK LATEST DATA
    # =====================
    current = data[data["QUARTER"] == latest_ssm_quarter]
    previous = data[data["QUARTER"] < latest_ssm_quarter].sort_values("DATE")

    has_latest = (not current.empty) and current["OBS.VALUE"].notna().any()

    # =====================
    # CREATE GRAPH (TEMP FILE)
    # =====================
    graph_path = f"{country}_temp.png"

    plt.figure()
    plt.plot(data["DATE"], data["OBS.VALUE"], label=country)
    plt.plot(ssm["DATE"], ssm["OBS.VALUE"], label="SSM")

    plt.gca().yaxis.set_major_formatter(mtick.PercentFormatter())

    quarters = data["DATE"].dt.to_period("Q").astype(str)

    plt.xticks(data["DATE"][::2], quarters[::2], rotation=90)

    plt.xlabel("Quarter")
    plt.ylabel("ROE (%)")
    plt.title(f"ROE - {country}")

    plt.legend()
    plt.tight_layout()

    plt.savefig(graph_path)
    plt.close()

    # =====================
    # TEXT BLOCK
    # =====================
    if not has_latest:

        text_block = f"""
        <p>
        Data is unavailable for the quarter ({latest_ssm_quarter}) 
        due to confidentiality issues.
        </p>
        """

    else:
        latest = current.iloc[-1]["OBS.VALUE"]

        prev = previous.iloc[-1]["OBS.VALUE"] if not previous.empty else None

        if prev is not None:
            trend = "increased" if latest > prev else "decreased"
        else:
            trend = "not available"

        ssm_value = ssm[ssm["QUARTER"] == latest_ssm_quarter].iloc[-1]["OBS.VALUE"]
        comparison = "above" if latest > ssm_value else "below"

        text_block = f"""
        <p>
        The ROE has <b>{trend}</b> compared to the previous quarter and is 
        currently <b>{comparison}</b> the SSM aggregate 
        (reference quarter {latest_ssm_quarter}).
        </p>
        """

    # =====================
    # CREATE EMAIL
    # =====================
    mail = outlook.CreateItem(0)

    mail.Subject = f"Profitability outlook - {country}"

    mail.HTMLBody = f"""
    <p>Dear colleague,</p>

    <p>
    As part of our quarterly profitability monitoring, please find below an analysis of the Return on Equity (ROE) of significant institutions in your jurisdiction, compared with the overall SSM aggregate.
    </p>

    <p>
    The data have been retrieved from the ECB Data Portal and are available at the following link:<br>
    <a href="https://data.ecb.europa.eu/data/datasets/SUP/SUP.Q.B01.W0._Z.I2003._T.SII._Z._Z._Z.PCT.C">
    ECB Supervisory Data – ROE
    </a>
    </p>

    {text_block}

    <p><b>ROE time series:</b></p>

    
    <p><img src="cid:myimage"></p>

    <p>Best regards,<br>
    """

    # =====================
    # EMBED IMAGE
    # =====================
    attachment = mail.Attachments.Add(os.path.abspath(graph_path))
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        "myimage"
    )

    # save as draft
    mail.Save()

    # optional: delete temp file
    os.remove(graph_path)

# =====================
# DONE
# =====================
print("Emails saved in Outlook Drafts!")
