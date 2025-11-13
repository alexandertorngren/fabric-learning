[StagingDefinition = [Kind = "FastCopy"]]
section Section1;
shared DimProduct = let
  Source = Excel.Workbook(Web.Contents("https://drakeanalytics-my.sharepoint.com/personal/alexander_torngren_drakeanalytics_se/Documents/Apps/Microsoft Power Query/Uploaded Files/FabricData.xlsx"), null, true),
  #"Navigation 1" = Source{[Item = "DimProduct", Kind = "Sheet"]}[Data],
  #"Promoted headers" = Table.PromoteHeaders(#"Navigation 1", [PromoteAllScalars = true]),
  #"Changed column type" = Table.TransformColumnTypes(#"Promoted headers", {{"ProductKey", Int64.Type}, {"ProductAlternateKey", type text}, {"ProductSubcategoryKey", type text}, {"WeightUnitMeasureCode", type text}, {"SizeUnitMeasureCode", type text}, {"EnglishProductName", type text}, {"StandardCost", type text}, {"FinishedGoodsFlag", Int64.Type}, {"Color", type text}, {"SafetyStockLevel", Int64.Type}, {"ReorderPoint", Int64.Type}, {"ListPrice", type text}, {"Size", type text}, {"SizeRange", type text}, {"DaysToManufacture", Int64.Type}, {"ProductLine", type text}, {"DealerPrice", type text}, {"Class", type text}, {"Style", type text}, {"ModelName", type text}, {"LargePhoto", type text}, {"EnglishDescription", type text}, {"FrenchDescription", type text}, {"ChineseDescription", type text}, {"ArabicDescription", type text}, {"HebrewDescription", type text}, {"ThaiDescription", type text}, {"GermanDescription", type text}, {"JapaneseDescription", type text}, {"TurkishDescription", type text}, {"StartDate", type datetime}, {"EndDate", type text}, {"Status", type text}, {"SpanishProductName", type text}, {"FrenchProductName", type text}, {"Weight", type number}})
in
  #"Changed column type";
shared DimProductCategory = let
  Source = Excel.Workbook(Web.Contents("https://drakeanalytics-my.sharepoint.com/personal/alexander_torngren_drakeanalytics_se/Documents/Apps/Microsoft Power Query/Uploaded Files/FabricData.xlsx"), null, true),
  #"Navigation 1" = Source{[Item = "DimProductCategory", Kind = "Sheet"]}[Data],
  #"Promoted headers" = Table.PromoteHeaders(#"Navigation 1", [PromoteAllScalars = true]),
  #"Changed column type" = Table.TransformColumnTypes(#"Promoted headers", {{"ProductCategoryKey", Int64.Type}, {"ProductCategoryAlternateKey", Int64.Type}, {"EnglishProductCategoryName", type text}, {"SpanishProductCategoryName", type text}, {"FrenchProductCategoryName", type text}})
in
  #"Changed column type";
shared DimProductSubcategory = let
  Source = Excel.Workbook(Web.Contents("https://drakeanalytics-my.sharepoint.com/personal/alexander_torngren_drakeanalytics_se/Documents/Apps/Microsoft Power Query/Uploaded Files/FabricData.xlsx"), null, true),
  #"Navigation 1" = Source{[Item = "DimProductSubcategory", Kind = "Sheet"]}[Data],
  #"Promoted headers" = Table.PromoteHeaders(#"Navigation 1", [PromoteAllScalars = true]),
  #"Changed column type" = Table.TransformColumnTypes(#"Promoted headers", {{"ProductSubcategoryKey", Int64.Type}, {"ProductSubcategoryAlternateKey", Int64.Type}, {"EnglishProductSubcategoryName", type text}, {"SpanishProductSubcategoryName", type text}, {"FrenchProductSubcategoryName", type text}, {"ProductCategoryKey", Int64.Type}})
in
  #"Changed column type";
shared FactInternetSales = let
  Source = Excel.Workbook(Web.Contents("https://drakeanalytics-my.sharepoint.com/personal/alexander_torngren_drakeanalytics_se/Documents/Apps/Microsoft Power Query/Uploaded Files/FabricData.xlsx"), null, true),
  #"Navigation 1" = Source{[Item = "FactInternetSales", Kind = "Sheet"]}[Data],
  #"Promoted headers" = Table.PromoteHeaders(#"Navigation 1", [PromoteAllScalars = true]),
  #"Changed column type" = Table.TransformColumnTypes(#"Promoted headers", {{"ProductKey", Int64.Type}, {"OrderDateKey", Int64.Type}, {"DueDateKey", Int64.Type}, {"ShipDateKey", Int64.Type}, {"CustomerKey", Int64.Type}, {"PromotionKey", Int64.Type}, {"CurrencyKey", Int64.Type}, {"SalesTerritoryKey", Int64.Type}, {"SalesOrderNumber", type text}, {"SalesOrderLineNumber", Int64.Type}, {"RevisionNumber", Int64.Type}, {"OrderQuantity", Int64.Type}, {"UnitPrice", type number}, {"ExtendedAmount", type number}, {"UnitPriceDiscountPct", Int64.Type}, {"DiscountAmount", Int64.Type}, {"ProductStandardCost", type number}, {"TotalProductCost", type number}, {"SalesAmount", type number}, {"TaxAmt", type number}, {"Freight", type number}, {"CarrierTrackingNumber", type text}, {"CustomerPONumber", type text}, {"OrderDate", type date}, {"DueDate", type date}, {"ShipDate", type date}})
in
  #"Changed column type";
