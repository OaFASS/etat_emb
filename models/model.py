# -*- coding: utf-8 -*-
from odoo import models, fields, api, _
from odoo.exceptions import AccessError, UserError, ValidationError
import xlwt
import xlsxwriter
from xlrd import open_workbook
import openpyxl
from datetime import timedelta
class GenerateFinancialAccountingStatetement(models.Model):
    _name = "financial.accounting.statement"
    company = fields.Many2one('res.company')
    date_deb = fields.Date(required=True, default="")
    date_fin = fields.Date(required=True, default="")
    
    def generate(self):
        file_path ='C:/Program Files/Odoo 16.0e.20230524/server/odoo/addons/ecf_module/static/src/classeur.xlsx'
        workbook = openpyxl.load_workbook(file_path)
        #FEUILLES DES ACTIFS 
        ## Frais de developpement
        sheet = workbook.worksheets[5]
        user = self.env.user
        company_id = user.company_id.id
        daten1 = self.date_deb - timedelta(days=365)
        daten2 = self.date_fin - timedelta (days=365)
    
        code_id1 = self.env['account.account'].search([('code', 'like', '211'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin),('parent_state','=','posted')])
        val1 = 0
        for record in values:
            val1 += record.credit
        sheet['D11'] = val1 #Brut 

        code_id2 = self.env['account.account'].search([('code', 'like', '2811'),('company_id', '=', company_id)])
        values = []

        for code in code_id2:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin),('parent_state','=','posted')])
        val2 = 0
        for record in values:
            val2 += record.credit       
        sheet['E11'] = val2 # Amortissement et dépréciations        

        values0 = []; values1 =[]
        for code in code_id1:
            values0 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2),('parent_state','=','posted')])
        for code in code_id2:
            values1 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2),('parent_state','=','posted')])
        val3 = 0 ; val4 = 0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val3 += record.credit
        sheet['G11'] = val3 - val4 # N-1

        ## Brevets, licences, logiciels, et  droits similaires
        code_id1 = self.env['account.account'].search([('code', 'like', '212'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin),('parent_state','=','posted')])
        
        code_id2 = self.env['account.account'].search([('code', 'like', '213'),('company_id', '=', company_id)])
        for code in code_id2:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin),('parent_state','=','posted')])
        
        code_id3 = self.env['account.account'].search([('code', 'like', '214'),('company_id', '=', company_id)])
        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin),('parent_state','=','posted')])

        val1 = 0
        for record in values:
            val1 += record.credit
        sheet['D12'] = val1 #Brut
        values = []
        code_id5 = self.env['account.account'].search([('code', 'like', '2812'),('company_id', '=', company_id)])
        code_id6 = self.env['account.account'].search([('code', 'like', '2813'),('company_id', '=', company_id)])
        code_id7 = self.env['account.account'].search([('code', 'like', '2814'),('company_id', '=', company_id)])
        for code in code_id5:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin),('parent_state','=','posted')])
        for code in code_id6:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin),('parent_state','=','posted')])
        for code in code_id7:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin),('parent_state','=','posted')])
        
        val2 = 0
        for record in values:
            val2 += record.credit
                
        sheet['E12'] = val2 #Amortissements et depreciations
        
        values0 = [] ; values1 =[]
        for code in code_id1:
            values0 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2),('parent_state','=','posted')])
        for code in code_id2:
            values0 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2),('parent_state','=','posted')])
        for code in code_id3:
            values0 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2),('parent_state','=','posted')])
        for code in code_id5:
            values1 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2),('parent_state','=','posted')])
        for code in code_id6:
            values1 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2),('parent_state','=','posted')])
        for code in code_id7:
            values1 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2),('parent_state','=','posted')])
        
        val3 = 0 ; val4 = 0
        for record in values:
            val3 += record.credit
        sheet['G12'] = val3 - val4

        #Fond commercial et droit de bail
      
        code_id1 = self.env['account.account'].search([('code', 'like', '216'),('company_id', '=', company_id)])
        code_id2 = self.env['account.account'].search([('code', 'like', '215'),('company_id', '=', company_id)])
        values = []
        
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin),('parent_state','=','posted')])
        for code in code_id2:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0           
        for record in values:
            val1 += record.credit
        sheet['D13'] = val1 # Brut
        
        code_id3 = self.env['account.account'].search([('code', 'like', '2915'),('company_id', '=', company_id)])
        code_id4 = self.env['account.account'].search([('code', 'like', '2916'),('company_id', '=', company_id)])
        
        values = []
        
        for code in code_id3:        
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        for code in code_id4:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val2 = 0           
        for record in values:
            val1 += record.credit
        sheet['E13'] = val1 #Amort et depreciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code_id1.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id2:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code_id2.id),('date','>=',daten1),('date','<=',daten2)])
        values0 += values1
        for code in code_id3:  
            values2 = self.env['account.move.line'].search([('account_id', '=', code_id3.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id4:
            values3 = self.env['account.move.line'].search([('account_id', '=', code_id4.id),('date','>=',daten1),('date','<=',daten2)])
        values2 += values3
        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values2:
            val4 += record.credit
        sheet['G13'] = val3 - val4 #Net N-1

        #Autres immobilisations incorporelles

        #Terrains
        code_id1 = self.env['account.account'].search([('code', 'like', '22'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0           
        for record in values:
            val1 += record.credit
        sheet['D16'] = val1 # Brut

        values = []
        code_id2 = self.env['account.account'].search([('code', 'like', '292'),('company_id', '=', company_id)])
        code_id3 = self.env['account.account'].search([('code', 'like', '282'),('company_id', '=', company_id)])
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        for code in code_id2:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0           
        for record in values:
            val1 += record.credit
        sheet['E16'] = val1 #amortissements et dépréciations

        values0 =[]
        for code in code_id1: 
            values0 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        values1 =[]
        for code in code_id2:   
            values1 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        values2 =[]
        for code in code_id3:  
            values2 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        values2 += values1
        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values2:
            val4 += record.credit
        sheet['G16'] = val3 - val4 #Net N-1

        #Batiments
        code_id1 = self.env['account.account'].search([('code', 'like', '231'),('company_id', '=', company_id)])
        code_id2 = self.env['account.account'].search([('code', 'like', '232'),('company_id', '=', company_id)])
        code_id1 += code_id2
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D18'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '2391'),('company_id', '=', company_id)])
        code_id4 = self.env['account.account'].search([('code', 'like', '2831'),('company_id', '=', company_id)])
        code_id3 += code_id4
        code_id5 = self.env['account.account'].search([('code', 'like', '2832'),('company_id', '=', company_id)])
        code_id3 += code_id5
        code_id6 = self.env['account.account'].search([('code', 'like', '2833'),('company_id', '=', company_id)])
        code_id7 = self.env['account.account'].search([('code', 'like', '2392'),('company_id', '=', company_id)])
        code_id8 = self.env['account.account'].search([('code', 'like', '2393'),('company_id', '=', company_id)])
        code_id3 += code_id6 + code_id7 + code_id8
        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit
        sheet['E18'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit
        sheet['G18'] = val3 - val4 #Net N-1

        # aménagements, agencement et installations

        code_id1 = self.env['account.account'].search([('code', 'like', '234'),('company_id', '=', company_id)])
        code_id2 = self.env['account.account'].search([('code', 'like', '235'),('company_id', '=', company_id)])
        code_id3 = self.env['account.account'].search([('code', 'like', '237'),('company_id', '=', company_id)])
        code_id4= self.env['account.account'].search([('code', 'like', '238'),('company_id', '=', company_id)])
        code_id1 += code_id2 + code_id3 + code_id4
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D20'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '2394'),('company_id', '=', company_id)])
        code_id4 = self.env['account.account'].search([('code', 'like', '2834'),('company_id', '=', company_id)])
        code_id3 += code_id4
        code_id5 = self.env['account.account'].search([('code', 'like', '2835'),('company_id', '=', company_id)])
        code_id3 += code_id5
        code_id6 = self.env['account.account'].search([('code', 'like', '2837'),('company_id', '=', company_id)])
        code_id7 = self.env['account.account'].search([('code', 'like', '2395'),('company_id', '=', company_id)])
        code_id8 = self.env['account.account'].search([('code', 'like', '2398'),('company_id', '=', company_id)])
        code_id9 = self.env['account.account'].search([('code', 'like', '2838'),('company_id', '=', company_id)])
        code_id3 += code_id6 + code_id7 + code_id8 + code_id9
        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit
        sheet['E20'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit
        sheet['G20'] = val3 - val4 #Net N-1

        # Matériel, mobilier et actifs biologiques

        code_id1 = self.env['account.account'].search([('code', 'like', '24'),('company_id', '=', company_id)])
        code_id2 = self.env['account.account'].search([('code', 'like', '245'),('company_id', '=', company_id)])
        code_id1 -= code_id2
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D21'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '284'),('company_id', '=', company_id)])
        code_id4 = self.env['account.account'].search([('code', 'like', '2845'),('company_id', '=', company_id)])
        code_id3 -= code_id4

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit
        sheet['E21'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit
        sheet['G21'] = val3 - val4 #Net N-1

        #Matériel de transport

        code_id1 = self.env['account.account'].search([('code', 'like', '245'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D22'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '2845'),('company_id', '=', company_id)])
        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit
        sheet['E22'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit
        sheet['G22'] = val3 - val4 #Net N-1

        #Avances et acomptes versés sur immobilisations

        code_id1 = self.env['account.account'].search([('code', 'like', '25'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D23'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '2951'),('company_id', '=', company_id)])
        code_id4 = self.env['account.account'].search([('code', 'like', '2952'),('company_id', '=', company_id)])
        code_id3 += code_id4

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit
        sheet['E23'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit
        sheet['G23'] = val3 - val4 #Net N-1

        # Titres de participation

        code_id1 = self.env['account.account'].search([('code', 'like', '26'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D25'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '296'),('company_id', '=', company_id)])
       

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E25'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit
        sheet['G25'] = val3 - val4 #Net N-1 

        #Autres immobilisations financières
        code_id1 = self.env['account.account'].search([('code', 'like', '27'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D25'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '297'),('company_id', '=', company_id)])
        

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E25'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit
        sheet['G25'] = val3 - val4 #Net N-1 

        #Actifs Circulants HAO

        code_id1 = self.env['account.account'].search([('code', 'like', '488'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D28'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '498'),('company_id', '=', company_id)])
        

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E28'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit
        sheet['G28'] = val3 - val4 #Net N-1 

        # Stock et en cours

        code_id1 = self.env['account.account'].search([('code', 'like', '31'),('company_id', '=', company_id)])
        code_id1 += self.env['account.account'].search([('code', 'like', '32'),('company_id', '=', company_id)])
        code_id1 += self.env['account.account'].search([('code', 'like', '33'),('company_id', '=', company_id)])
        code_id1 += self.env['account.account'].search([('code', 'like', '34'),('company_id', '=', company_id)])
        code_id1 += self.env['account.account'].search([('code', 'like', '35'),('company_id', '=', company_id)])
        code_id1 += self.env['account.account'].search([('code', 'like', '37'),('company_id', '=', company_id)])
        code_id1 += self.env['account.account'].search([('code', 'like', '38'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D29'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '391'),('company_id', '=', company_id)])

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E29'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit
        sheet['G29'] = val3 - val4 #Net N-1

        #Creances et emplois assimilés

        code_id1 = self.env['account.account'].search([('code', 'like', '41'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D30'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '491'),('company_id', '=', company_id)])
        code_id3 += self.env['account.account'].search([('code', 'like', '492'),('company_id', '=', company_id)])
        code_id3 += self.env['account.account'].search([('code', 'like', '493'),('company_id', '=', company_id)])
        code_id3 += self.env['account.account'].search([('code', 'like', '494'),('company_id', '=', company_id)])
        code_id3 += self.env['account.account'].search([('code', 'like', '495'),('company_id', '=', company_id)])

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E30'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit
        sheet['G30'] = val3 - val4 #Net N-1

        # Fournisseurs Avances versées

        code_id1 = self.env['account.account'].search([('code', 'like', '409'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D31'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '489'),('company_id', '=', company_id)])

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E31'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit + record.debit
        sheet['G31'] = val3 - val4 #Net N-1

        #clients 

        code_id1 = self.env['account.account'].search([('code', 'like', '41'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D32'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '491'),('company_id', '=', company_id)])

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E32'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit + record.debit
        sheet['G32'] = val3 - val4 #Net N-1

        # Autres créances 

        code_id1 = self.env['account.account'].search([('code', 'like', '47'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D33'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '479'),('company_id', '=', company_id)])

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E33'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit + record.debit
        sheet['G33'] = val3 - val4 #Net N-1

        # Tritres de placement
        code_id1 = self.env['account.account'].search([('code', 'like', '50'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D35'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '590'),('company_id', '=', company_id)])

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E35'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit + record.debit
        sheet['G35'] = val3 - val4 #Net N-1

        #Valeurs à encaisser
        code_id1 = self.env['account.account'].search([('code', 'like', '51'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D36'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '591'),('company_id', '=', company_id)])

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E36'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit + record.debit
        sheet['G36'] = val3 - val4 #Net N-1
        
        #Banques, chèques postaux, caisse et assimilés

        code_id1 = self.env['account.account'].search([('code', 'like', '51'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D37'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '591'),('company_id', '=', company_id)])

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E37'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit + record.debit
        sheet['G37'] = val3 - val4 #Net N-1

        #Ecart de conversion-Actif

        code_id1 = self.env['account.account'].search([('code', 'like', '518'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D39'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '499'),('company_id', '=', company_id)])

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E39'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit + record.debit
        sheet['G39'] = val3 - val4 #Net N-1

        #FEUILLES DES PASSIFS
        
        #Capital
        code_id1 = self.env['account.account'].search([('code', 'like', '518'),('company_id', '=', company_id)])
        values = []
        for code in code_id1:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
        val1 = 0          
        for record in values:
            val1 += record.credit
        sheet['D39'] = val1 # Brut

        values = []

        code_id3 = self.env['account.account'].search([('code', 'like', '499'),('company_id', '=', company_id)])

        for code in code_id3:
            values += self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',self.date_deb),('date','<=',self.date_fin)])
    
        val1 = 0           
        for record in values:
            val1 += record.credit + record.debit
        sheet['E39'] = val1 #amortissements et dépréciations

        for code in code_id1: 
            values0 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])
        for code in code_id3:   
            values1 = self.env['account.move.line'].search([('account_id', '=', code.id),('date','>=',daten1),('date','<=',daten2)])

        val3 = 0 ; val4 =0
        for record in values0:
            val3 += record.credit
        for record in values1:
            val4 += record.credit + record.debit
        sheet['G39'] = val3 - val4 #Net N-1



        
        

        workbook.save('C:/Program Files/Odoo 16.0e.20230524/server/odoo/addons/ecf_module/static/src/classeur.xlsx')
        workbook.close()
        
        return {
                'type': 'ir.actions.act_url',
                'url': '/ecf_module/static/src/classeur.xlsx',
                'target': 'target',
        }