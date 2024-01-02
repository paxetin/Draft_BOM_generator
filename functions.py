import requests
import json
import pandas as pd
import numpy as np
from tqdm import tqdm

class Setting:
     def __init__(self) -> None:
         pass

class Searcher:
    def __init__(self, by: str='keyword') -> None:
        """
        Attributes:
            by : currently support searching by 
                1. part number as 'partnumber' 
                2. keyword as 'keyword'
        """
        self.by = by
        self.url = "https://api.mouser.com/api/"
        version = "v1"
        self.endpoint = f"{version}/search/{by}"

        # This part is used to set the designated format such as json.
        self.header = {
            'accept': 'application/json',
            'Content-Type': 'application/json'
        }

        with open('./config.json', 'r') as f:
            config = json.load(f)
        
        self.params = {
            'apikey': config.get("MOUSER_API_KEY")
        }

    def __make_request(self) -> dict:
        response = requests.post(self.url + self.endpoint,
                                 headers=self.header,
                                 params=self.params,
                                 json=self.data)
         
        try:
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError:
            raise Exception(f'Error: {response.status_code} - {response.text}')
        
    def get(self, value) -> dict:
        if self.by.lower() == 'partnumber':
            self.data = {
                'SearchByPartRequest': {
                    'mouserPartNumber': value
                }
            }
        elif self.by.lower() == 'keyword':
            self.data = {
                "SearchByKeywordRequest": {
                    "keyword": value,
                    "records": 0,
                    "startingRecord": 0,
                    "searchOptions": "",
                    "searchWithYourSignUpLanguage": ""
                }
            }
        return self.__make_request()
    

class BOM_cost_Generator:
    def __init__(self) -> None:
        self.BOM = pd.DataFrame(columns=['MPN', 'CPN', 'Description', 'Category', 'Manufacturer', 'ManufacturerPartNumber', 'L/T', 'Unit Price', 'Availability', 'Packaging'])
        self.idx = 0

    def __decompose(self, data, target_col):
        pass

    def extract(self, df):
        min_unit_price = np.min(list(map(lambda x: float(x.strip("$")), pd.DataFrame(df['PriceBreaks'][0])['Price'].tolist())))
        return [df['Description'], df['Category'], df['Manufacturer'], df['ManufacturerPartNumber'], df['LeadTime'], min_unit_price, df['AvailabilityInStock']]

    def append(self, df: pd.DataFrame, MPN: str='TBD', CPN: str='TBD') -> None:
        if len(df) > 1:
            for i, row in enumerate(df.iterrows()):
                self.BOM.loc[self.idx, :] = [MPN, CPN] + self.extract()
                self.idx += 1
        else:
            self.BOM.loc[self.idx, :] = [MPN, CPN] + self.extract()
            self.idx += 1


class BOMGenerator:
    def __init__(self) -> None:
        self.__BOM = pd.DataFrame(columns=['#', 'Vendor', 'Type', 'P/N', 'Specification', 'Quantity', 'Number', 'FPN', 'Package'])
        self.idx = 0

    def __extract(self, df):
        package = df['ProductAttributes'][0][ 'AttributeValue'] if len(df['ProductAttributes']) > 0 else None
        return [self.idx+1, df['Manufacturer'], df['Category'], df['ManufacturerPartNumber'], df['Description'], 
                None, None, None, 
                package]

    def append(self, df: pd.DataFrame, part, specify: int=0):
        if len(df) == 0: 
            self.__BOM.loc[self.idx, :] = [self.idx+1, None, None, None, part, None, None, None, None]
        else:
            self.__BOM.loc[self.idx, :] = self.__extract(df.iloc[specify])
        self.idx += 1

    @property
    def BOM(self):
        return self.__BOM

    def to_xlxs(self, filename):
        with pd.ExcelWriter(f"{filename}.xlsx") as writer:
            for idx, row in enumerate(self.BOM.iterrows()):
                row.to_excel(writer, sheet_name=str(idx), index=False)


def xlsx_loader(path, header=1, sheet_name='OrCAD BOM list'):
    return pd.read_excel(io=path, 
                         header=header,
                         sheet_name=sheet_name)['Part'].dropna()


def desc_parser(part_name):
    spliter = [m for m in list(set(str(part_name))) if m in ['/', ',', '_']]
    if len(spliter) > 0:
        return part_name.replace(spliter[0], ';')
    else:
        return part_name



def _run(target_file=None, 
        designated_name="test123"):
    """Attributes:
        target: The original BOM file path as input.
        designated_name: The file name for generated BOM.
    """

    input_bom = xlsx_loader(target_file)

    try:
        parts = input_bom.tolist()
    except:
        raise KeyError('"Part" is not found in axis.')

    searcher = Searcher(by = 'keyword')
    generator = BOMGenerator()

    with pd.ExcelWriter(f"./{designated_name}.xlsx") as writer:

        # init to keep the draft BOM at the first page.
        generator.BOM.to_excel(writer, sheet_name='draft BOM')
            
        for idx, part in enumerate(tqdm(parts)):
            data = searcher.get(desc_parser(part))
            df = pd.json_normalize(data['SearchResults']['Parts'])
            df.to_excel(writer, sheet_name=str(idx+1), index=False)

            generator.append(df, part)

        # update the draft BOM per row.
        generator.BOM.to_excel(writer, sheet_name='draft BOM', index=False)