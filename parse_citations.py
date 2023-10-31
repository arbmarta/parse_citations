import pandas as pd

def parse_citation(citation_text):
    lines = citation_text.strip().split('\n')
    authors = lines[0].strip()
    title = lines[1].strip()
    journal = "Landscape and Urban Planning"
    
    volume_issue = lines[3].replace("Volume ", "").replace("Issues ", "").replace(",", "")
    volume, issue = volume_issue.split() if " " in volume_issue else (volume_issue, "")
    
    year = lines[4].strip()
    year = ''.join(filter(str.isdigit, year))  # Extracting only the numeric part for the year
    
    doi = lines[7].replace("https://doi.org/", "").strip()
    return authors, title, year, journal, volume, issue, doi

def main():
    input_file_path = 'citations.txt'
    output_file_path = 'output.xlsx'
    
    try:
        with open(input_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError as e:
        print("Error reading the file. Please make sure the file is saved with UTF-8 encoding.")
        print("Technical details:", e)
        return
    except Exception as e:
        print("An error occurred:", e)
        return

    citations = content.split('\n\n')
    data = []
    for citation in citations:
        if citation.strip():
            try:
                parsed_citation = parse_citation(citation)
                data.append(parsed_citation)
            except Exception as e:
                print("Error parsing citation:", e)
                print("Citation text:", citation)
                continue

    df = pd.DataFrame(data, columns=['Authors', 'Title', 'Year', 'Journal', 'Volume', 'Issue', 'DOI'])
    print(df)

    # Save to Excel file
    df.to_excel(output_file_path, index=False)
    print(f'Data saved to {output_file_path}')

if __name__ == "__main__":
    main()
