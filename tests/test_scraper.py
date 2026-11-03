import io
import json
import unittest
import zipfile
from pathlib import Path
from scraper import Template, parse_product, classify, money, NS, discover, bardahl_package, phone_package, build_row
from lxml import etree


class ScraperTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.template = Template(Path(__file__).resolve().parents[1]/"input/template.xlsx")

    def test_original_gallery_and_sku(self):
        product={"@type":"Product","name":"Кабел Phone Planet","sku":"00-TEST","offers":{"price":"5.044","priceCurrency":"EUR","availability":"https://schema.org/InStock"}}
        source='<html><script type="application/ld+json">'+json.dumps(product)+'</script><a id="product-1-main-image-anchor" href="/storage/app/images/products/main/test.webp"><img src="/thumbs/test.webp"></a></html>'
        p=parse_product(source,"https://accessbg.com/product/test",{"phone_planet"})
        self.assertEqual(p["sku"],"00-TEST")
        self.assertEqual(p["price"],5.04)
        self.assertEqual(p["images"],["https://accessbg.com/storage/app/images/products/main/test.webp"])
        self.assertEqual(p["availability"],"in_stock")

    def test_category_edge_cases(self):
        def p(name,group):
            return {"name":name,"groups":[group],"description":""}
        self.assertEqual(classify(p("Macлo ВАRDАНL ХТС 5W40 1L","bardahl")),"20416")
        self.assertEqual(classify(p("Масло BARDAHL XTG 75W90 1L","bardahl")),"20414")
        self.assertEqual(classify(p("XBF DOT 5.1 450ml","bardahl")),"20373")
        self.assertIsNone(classify(p("Препарат BARDAHL Полираща паста 500ml","bardahl")))
        self.assertIsNone(classify(p("6 в 1 Phone Planet SET за Iphone","phone_planet")))
        self.assertEqual(classify(p("Видеорегистратор Phone Planet DC01","phone_planet")),"22906")

    def test_required_and_condition_rules(self):
        errors=self.template.validate({"E":"24321","DS":"USB","DZ":"Without Battery"})
        self.assertTrue(any("Operating Voltage" in e for e in errors))
        self.assertFalse(any("Plug Type" in e for e in errors))
        self.assertTrue(any("Country/Region of Origin" in e for e in errors))
        errors=self.template.validate({"E":"20364","OV":"NOT-A-COUNTRY"})
        self.assertTrue(any("Invalid dropdown value OV" in e for e in errors))

    def test_template_package_preservation(self):
        # Test in memory: all unrelated package members must remain byte-identical.
        data=self.template.workbook_bytes([{"E":"20364","N":"00-SKU","L":"=literal"}])
        with zipfile.ZipFile(self.template.path) as original,zipfile.ZipFile(io.BytesIO(data)) as result:
            self.assertEqual(original.namelist(),result.namelist())
            target=self.template.paths["Template"]
            for name in original.namelist():
                if name!=target:
                    self.assertEqual(original.read(name),result.read(name),name)
            before=etree.fromstring(original.read(target))
            after=etree.fromstring(result.read(target))
            for n in range(1,5):
                xp=f's:sheetData/s:row[@r="{n}"]'
                self.assertEqual(etree.tostring(before.find(xp,NS)),etree.tostring(after.find(xp,NS)))
            cells=after.findall('s:sheetData/s:row[@r="5"]/s:c',NS)
            self.assertEqual(next(c for c in cells if c.get("r")=="N5").get("t"),"inlineStr")
            self.assertIsNone(next(c for c in cells if c.get("r")=="L5").find('s:f',NS))
            self.assertIsNotNone(after.find('s:dataValidations',NS))

    def test_rounding(self):
        self.assertEqual(money("5.045"),5.05)
        self.assertEqual(money("€ 7,50"),7.5)

    def test_bardahl_user_weight_rule_and_model_numbers(self):
        for name,weight in [('Масло 1L',1100),('Масло 5L',5500),('Добавка 300ml',330),('PLASMA 01L BOOSTER 400ml',440),('Спрей 125m1',138),('Препарат 150gr',165),('Добавка - 1 Литър',1100)]:
            result,method=bardahl_package({'name':name,'description':'','sku':'TEST'})
            self.assertEqual(result['OE'],weight,name)
            self.assertTrue(all(result[c]>0 for c in ['OF','OG','OH']))
        # Dosing instructions must never be interpreted as package capacity.
        result,method=bardahl_package({'name':'Добавка','description':'Сипете в 30 литра гориво','sku':'TEST'})
        self.assertNotIn('OM',result)
        self.assertTrue(method['basis'].startswith('estimated'))

    def test_bardahl_eu_rp_blank_and_phone_traits(self):
        p={'sku':'TEST','name':'Масло 1L','groups':['bardahl'],'description':'','price':10,'regular_price':10,'images':['https://accessbg.com/storage/app/images/products/main/test.webp'],'issues':[],'currency':'EUR','availability':'in_stock'}
        config={'defaults':{'eu_responsible_person':'I M G Ltd.'},'bardahl':{'estimate_packages':True,'omit_eu_responsible_person':True}}
        row,issues=build_row(p,config,self.template,{'TEST':{'QX':'I M G Ltd.'}})
        self.assertNotIn('QX',row)
        self.assertEqual(row['OE'],1100)
        p.update({'name':'Кабел Phone Planet USB-C 60W','groups':['phone_planet'],'description':'USB Type-C 5V'})
        row,issues=build_row(p,{'phone_planet':{'estimate_packages':True}},self.template,{})
        self.assertEqual(row['EZ'],'60W')
        self.assertEqual(row['DT'],'5V')
        self.assertNotIn('EY',row) # output rating is not proof of input power

    def test_paginated_scope_and_coverage(self):
        from urllib.parse import urlparse,parse_qs
        class FakeClient:
            def get(self,url,fresh=False):
                if '/getFilteredProducts?' not in url:
                    return '<input id="selected-category-id" value="118">'
                q=parse_qs(urlparse(url).query)
                group='phone' if 'searchTerm' in q else 'bardahl'
                page=int(q['page'][0])
                item=f'https://accessbg.com/product/{group}-{page}'
                return json.dumps({'success':True,'data':{'html':f'<div class="product-card"><a href="{item}">X</a></div><a href="https://accessbg.com/product/RELATED">Ignore</a>',
                  'paginationMetaData':{'totalProducts':2,'lastPage':2}}})
        urls,counts=discover(FakeClient())
        self.assertEqual(len(urls),4)
        self.assertFalse(any('RELATED' in url for url in urls))
        self.assertEqual(counts['bardahl']['discovered'],2)


if __name__=="__main__":
    unittest.main()
