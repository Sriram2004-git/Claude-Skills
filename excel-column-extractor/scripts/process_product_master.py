import pandas as pd
import sys
import os
from datetime import datetime

def main(source_path, price_list_path, dest_folder):
    try:
        if not os.path.exists(source_path):
            print(f"Error: Source file '{source_path}' does not exist.")
            return False
            
        if not os.path.exists(price_list_path):
            print(f"Error: Price list file '{price_list_path}' does not exist.")
            return False
            
        if not os.path.exists(dest_folder):
            print(f"Creating destination folder '{dest_folder}'...")
            os.makedirs(dest_folder, exist_ok=True)

        # 1. Read Product Master, only extracting specific columns
        print(f"Reading source file (extracting specific columns): {os.path.basename(source_path)}...")
        required_cols = ["Material ID", "Ordering Code Packed", "MPG Product Description"]
        df_source = pd.read_excel(source_path, usecols=required_cols)
        total_rows = len(df_source)
        print(f"Extracted columns: {required_cols} from {total_rows} rows.")

        # 2. Read Price List
        print(f"Reading price list: {os.path.basename(price_list_path)}...")
        df_price = pd.read_excel(price_list_path)
        
        pl_fields = ["Material_NR", "HSN", "Techslice Price", "Market_List_Price_Value"]
        missing_pl_cols = [col for col in pl_fields if col not in df_price.columns]
        if missing_pl_cols:
            print(f"Error: Required columns missing from Price List: {', '.join(missing_pl_cols)}")
            return False
            
        df_price_filtered = df_price[pl_fields].copy()

        # 3. Match Material ID with Material_NR
        print("Matching Material ID with Material_NR in-memory...")
        enriched_df = pd.merge(
            df_source, 
            df_price_filtered, 
            left_on="Material ID", 
            right_on="Material_NR", 
            how="left"
        )

        # 4. Transform into final template structure
        print("Transforming data into final template structure...")
        final_columns = [
            "Item Name", "SKU", "HSN/SAC", "Description", "Rate", "MRP",
            "Product Type", "Account", "Usage unit", "Purchase Description", 
            "Purchase Rate", "Item Type", "Purchase Account", "Inventory Account", 
            "Reorder Point", "Vendor", "Vendor Number", "Initial Stock", "Initial Stock Rate", 
            "Stock On Hand", "Status", "Taxability Type", "Exemption Reason", 
            "Inter State Tax Name", "Inter State Tax Type", "Inter State Tax Rate", 
            "Intra State Tax Name", "Intra State Tax Type", "Intra State Tax Rate", 
            "Warehouse Name", "CF.custom_field"
        ]
        
        final_df = pd.DataFrame(columns=final_columns)
        
        final_df["Item Name"] = enriched_df["Ordering Code Packed"]
        final_df["SKU"] = enriched_df["Material ID"]
        final_df["HSN/SAC"] = enriched_df["HSN"]
        final_df["Description"] = enriched_df["MPG Product Description"]
        final_df["Rate"] = enriched_df["Market_List_Price_Value"]
        final_df["MRP"] = enriched_df["Market_List_Price_Value"]
        final_df["Purchase Description"] = enriched_df["MPG Product Description"]
        final_df["Purchase Rate"] = enriched_df["Techslice Price"]
        
        # Default values
        final_df["Product Type"] = "goods"
        final_df["Account"] = "Sales"
        final_df["Usage unit"] = "pcs"
        final_df["Item Type"] = "Sales"
        final_df["Purchase Account"] = "Cost of Goods Sold"
        final_df["Inventory Account"] = "Inventory Asset"
        final_df["Vendor"] = "SANDVIK COROMANT INDIA PRIVATE LIMITED"
        final_df["Vendor Number"] = 12025001
        final_df["Status"] = "Active"
        final_df["Inter State Tax Name"] = "IGST18"
        final_df["Inter State Tax Type"] = "Simple"
        final_df["Inter State Tax Rate"] = 18
        final_df["Intra State Tax Name"] = "GST18"
        final_df["Intra State Tax Type"] = "Group"
        final_df["Intra State Tax Rate"] = 18

        total_records = len(final_df)
        
        # Calculate split files
        if total_records <= 25000:
            num_files = 1
        else:
            num_files = max(2, int(total_records / 25000 + 0.5))
            
        month_year = datetime.now().strftime("%B-%Y")
        
        dest_paths = []
        if num_files == 1:
            base_name = f"Product_Master_Final_{month_year}.xlsx"
            dest_path = os.path.join(dest_folder, base_name)
            counter = 1
            while os.path.exists(dest_path):
                name_no_ext = f"Product_Master_Final_{month_year}_{counter}"
                dest_path = os.path.join(dest_folder, f"{name_no_ext}.xlsx")
                counter += 1
            dest_paths.append(dest_path)
        else:
            base_exists = False
            for n in range(1, num_files + 1):
                path = os.path.join(dest_folder, f"Product_Master_Final_Part-{n}-{month_year}.xlsx")
                if os.path.exists(path):
                    base_exists = True
                    break
            
            if not base_exists:
                for n in range(1, num_files + 1):
                    dest_paths.append(os.path.join(dest_folder, f"Product_Master_Final_Part-{n}-{month_year}.xlsx"))
            else:
                counter = 1
                while True:
                    any_exists = False
                    temp_paths = []
                    for n in range(1, num_files + 1):
                        path = os.path.join(dest_folder, f"Product_Master_Final_Part-{n}-{month_year}_{counter}.xlsx")
                        temp_paths.append(path)
                        if os.path.exists(path):
                            any_exists = True
                            break
                    if not any_exists:
                        dest_paths = temp_paths
                        break
                    counter += 1
                    
        # Save files
        print(f"Splitting data into {num_files} files...")
        for i in range(num_files):
            start_idx = i * 25000
            if i == num_files - 1:
                chunk_df = final_df.iloc[start_idx:]
            else:
                end_idx = (i + 1) * 25000
                chunk_df = final_df.iloc[start_idx:end_idx]
                
            dest_path = dest_paths[i]
            print(f"Saving Part {i+1}/{num_files} ({len(chunk_df)} records) to: {dest_path}...")
            chunk_df.to_excel(dest_path, index=False)

        print("\n--- Transformation Summary ---")
        print(f"Total Records: {total_records}")
        print(f"Total Files Generated: {num_files}")
        for idx, path in enumerate(dest_paths):
            print(f"  - File {idx+1}: {os.path.basename(path)}")
        print(f"Pricing Matches (MRP): {enriched_df['Market_List_Price_Value'].notnull().sum()}")
        print(f"Status: Success")
        print("------------------------------")
        return True

    except Exception as e:
        print(f"\nError: An unexpected error occurred: {str(e)}")
        return False

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python process_product_master.py <source_path> <price_list_path> <dest_folder>")
        sys.exit(1)
        
    src = sys.argv[1]
    price = sys.argv[2]
    dst = sys.argv[3]
    
    if main(src, price, dst):
        sys.exit(0)
    else:
        sys.exit(1)
