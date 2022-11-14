print("ayushi")
import pandas as pd
import openpyxl as xl
company_orderReport = "D:\PROJECT\FIles\Company X - Order Report.xlsx";
courier_companyInvoice = "D:\PROJECT\FIles\Courier Company - Invoice.xlsx";
company_PincodeZones = "D:\PROJECT\FIles\Company X - Pincode Zones.xlsx";
Company_SKUMaster = "D:\PROJECT\FIles\Company X - SKU Master.xlsx";
courier_CompanyRates = "D:\PROJECT\FIles\Courier Company - Rates.xlsx";
expected_Result = "D:\PROJECT\FIles\Expected_Result.xlsx";
company_orderReport_DF = pd.read_excel(company_orderReport);
courier_companyInvoice_DF = pd.read_excel(courier_companyInvoice);
company_SKUMaster_DF = pd.read_excel(Company_SKUMaster);
company_PincodeZones_DF = pd.read_excel(company_PincodeZones);
courier_CompanyRates_DF = pd.read_excel(courier_CompanyRates);
externOrderNo_List = courier_companyInvoice_DF['Order ID'].tolist();

columns = ['Order ID','AWB Number','Total weight as per X (KG)','Weight slab as per X (KG)','Total weight as per Courier Company (KG)',
           'Weight slab charged by Courier Company (KG)','Delivery Zone as per X','Delivery Zone charged by Courier Company',
           'Expected Charge as per X (Rs.)','Charges Billed by Courier Company (Rs.)','Difference Between Expected Charges and Billed Charges (Rs.)'
          ]
totalValuesList = [];

columns2 = ['Orders Summary','Count','Amount (Rs.)'];
totalValuesList2 = [];

correctly_charged = ['Total orders where X has been correctly charged'];
overcharged = ['Total Orders where X has been overcharged'];
undercharged = ['Total Orders where X has been undercharged'];

count1=0;
count2=0;
count3=0;

amount1=0;
amount2=0;
amount3=0;


for orderId in externOrderNo_List:
    valuesList = [];
    valuesList.append(orderId);
    
    filtered_courier_companyInvoice_DF = courier_companyInvoice_DF[courier_companyInvoice_DF['Order ID'] == orderId];
    AWB_Code = int(filtered_courier_companyInvoice_DF['AWB Code']);
    valuesList.append(AWB_Code);

    filterd_company_orderReport_DF = company_orderReport_DF[company_orderReport_DF['ExternOrderNo']==orderId];
    SKU_NumberList = filterd_company_orderReport_DF['SKU'];

    Total_weight_as_per_X= 0;
    for sku_number in SKU_NumberList:
        filterd_order_QtyList = filterd_company_orderReport_DF[filterd_company_orderReport_DF['SKU']==sku_number];
        order_Qty = filterd_order_QtyList['Order Qty'];
        filtered_company_SKUMaster_DF = company_SKUMaster_DF[company_SKUMaster_DF['SKU']==sku_number]
        wieghtPerSingleItem = filtered_company_SKUMaster_DF['Weight (g)'];
        wieghtPerSingleItem = wieghtPerSingleItem.tolist()[0];
        Total_weight_as_per_X+=float(order_Qty.tolist()[0])*float(wieghtPerSingleItem);
    Total_weight_as_per_X = Total_weight_as_per_X/1000;
    valuesList.append(Total_weight_as_per_X);
    
    temp_Weight_slab_as_per_X = str(Total_weight_as_per_X).split('.');
    Total_slab_weight_as_per_X = float(temp_Weight_slab_as_per_X[0]);
    temp_slab_weightB = float(temp_Weight_slab_as_per_X[1][0]);
    if(temp_slab_weightB<=0.0):
        pass;
    elif(temp_slab_weightB<5.0):
        Total_slab_weight_as_per_X += 0.5;
    elif(temp_slab_weightB>5.0):
        Total_slab_weight_as_per_X += 1; 
    valuesList.append(Total_slab_weight_as_per_X);
    
    Total_weight_as_per_Courier_Company = float(filtered_courier_companyInvoice_DF['Charged Weight']);
    valuesList.append(Total_weight_as_per_Courier_Company);
    
    temp_Total_weight_as_per_Courier_Company = str(Total_weight_as_per_Courier_Company).split('.');
    Total_slab_weight_as_per_Courier_Company = float(temp_Total_weight_as_per_Courier_Company[0]);
    temp_slab_weightB = float(temp_Total_weight_as_per_Courier_Company[1][0]);
    if(temp_slab_weightB<=0.0):
        pass;
    elif(temp_slab_weightB<5.0):
        Total_slab_weight_as_per_Courier_Company += 0.5;
    elif(temp_slab_weightB>5.0):
        Total_slab_weight_as_per_Courier_Company += 1; 
    valuesList.append(Total_slab_weight_as_per_Courier_Company);

    customer_Pin_Code = int(filtered_courier_companyInvoice_DF['Customer Pincode']);
    filtered_company_PincodeZones_DF = company_PincodeZones_DF[company_PincodeZones_DF['Customer Pincode']==customer_Pin_Code]
    Delivery_Zone_as_per_X = filtered_company_PincodeZones_DF['Zone'];
    valuesList.append(Delivery_Zone_as_per_X.tolist()[0].upper());
    
    Delivery_Zone_charged_by_Courier_Company = filtered_courier_companyInvoice_DF['Zone'];
    valuesList.append(Delivery_Zone_charged_by_Courier_Company.tolist()[0].upper());
    
    Type_of_Shipment = str(filtered_courier_companyInvoice_DF['Type of Shipment']);
    company_fixed_charge = "fwd_"+Delivery_Zone_as_per_X+"_fixed";
    company_additional_charge = "fwd_"+Delivery_Zone_as_per_X+"_additional";
    if(Type_of_Shipment == "Forward charges"):
        company_fixed_charge = "fwd_"+Delivery_Zone_as_per_X+"_fixed";
        company_additional_charge = "fwd_"+Delivery_Zone_as_per_X+"_additional";
    elif(Type_of_Shipment == "Forward and RTO charges"):
        company_fixed_charge = "rto_"+Delivery_Zone_as_per_X+"_fixed";
        company_additional_charge = "rto_"+Delivery_Zone_as_per_X+"_additional";
    
    no_of_multiple = Total_slab_weight_as_per_X / 0.5;
    temp_company_fixed_charge = courier_CompanyRates_DF[company_fixed_charge].values[0];
    temp_company_additional_charge = courier_CompanyRates_DF[company_additional_charge].values[0];
    Expected_Charge_as_per_X = temp_company_fixed_charge + (float(no_of_multiple-1)*temp_company_additional_charge);
    Expected_Charge_as_per_X = round(Expected_Charge_as_per_X[0],1);
    valuesList.append(Expected_Charge_as_per_X);

    Charges_Billed_by_Courier_Company = round(filtered_courier_companyInvoice_DF['Billing Amount (Rs.)'],1);
    valuesList.append(Charges_Billed_by_Courier_Company.tolist()[0]);

    Difference_Between_Expected_Charges_and_Billed_Charges = Expected_Charge_as_per_X - Charges_Billed_by_Courier_Company.tolist()[0];
    valuesList.append(Difference_Between_Expected_Charges_and_Billed_Charges);
    
    totalValuesList.append(valuesList);
    
    if(Difference_Between_Expected_Charges_and_Billed_Charges==0):
        count1+=1;
        amount1+=Charges_Billed_by_Courier_Company.tolist()[0];
    elif(Difference_Between_Expected_Charges_and_Billed_Charges<0):
        count2+=1;
        amount2+=Charges_Billed_by_Courier_Company.tolist()[0];
    elif(Difference_Between_Expected_Charges_and_Billed_Charges>0):
        count3+=1;
        amount3+=Charges_Billed_by_Courier_Company.tolist()[0];
    

df1 = pd.DataFrame(totalValuesList,
                   columns=columns)                     
df1.to_excel("Expected_Result_final.xlsx") 

correctly_charged.append(count1);
correctly_charged.append(amount1);
overcharged.append(count2);
overcharged.append(amount2);
undercharged.append(count3);
undercharged.append(amount3);

df2 = pd.DataFrame([correctly_charged,overcharged,undercharged],
                   columns=columns2)
df2.to_excel("Expected_Summary_Result.xlsx")
    
