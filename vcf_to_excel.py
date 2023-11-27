import vcf
import pandas as pd

def parse_vcf(vcf_path):
    vcf_reader = vcf.Reader(filename=vcf_path)
    records = []

    for record in vcf_reader:
        chrom = record.CHROM
        pos = record.POS
        ref = str(record.REF)
        alt = ",".join(str(a) for a in record.ALT)
        qual = record.QUAL
        info = record.INFO

        # Extract information from each sample
        samples_data = {}
        for sample in record.samples:
            sample_data = {
                'GT': sample['GT'],
                'PL': ":".join(str(p) for p in sample.data.PL) if 'PL' in sample.data else None,
                # Add more fields as needed
            }
            samples_data[sample.sample] = sample_data

        records.append([chrom, pos, ref, alt, qual, info, samples_data])

    columns = ['Chromosome', 'Position', 'Reference', 'Alternate', 'Quality', 'Info', 'Samples']
    df = pd.DataFrame(records, columns=columns)
    return df

def convert_to_excel(vcf_path, excel_path):
    df = parse_vcf(vcf_path)
    df.to_excel(excel_path, index=False)
    print(f"Conversion completed. Excel file saved at: {excel_path}")

# Replace 'your_variant_file.vcf' with the path to your genome comparison VCF file
vcf_file_path = 'Edwardsiella_Sample+Reference_Merge_YJ.vcf'

# Replace 'output_excel_file.xlsx' with the desired output Excel file path
output_excel_path = 'output_excel_file4.xlsx'

convert_to_excel(vcf_file_path, output_excel_path)
