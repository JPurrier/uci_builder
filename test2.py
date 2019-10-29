import pandas as pd
from xml.etree.ElementTree import Element, SubElement, Comment, tostring
import datetime
from xml.dom import minidom
from xml.etree import ElementTree
from excelimport import ExcelImporter


class Buildetwork(object):

    def __init__(self):
        self.my_filetypes = [('Excel File', 'xlsx'), ('Excel file', 'xlx')]
        self.sheet_bridge_domain = 'bridge_domain'
        self.sheet_bd_subnet = 'bd_subnet'
        self.sheet_bd_l3out = 'bd_l3out'
        self.sheet_contract = 'contract'
        self.sheet_epg = 'end_point_group'
        self.sheet_epg_subnet = 'epg_subnet'
        self.sheet_epg_domain_ass = 'epg_domain_association'
        self.sheet_stat_bnd = 'epg_static_binding'
        self.sheet_app_prof = 'application_profile'
        self.sheet_tenant = 'tenant'
        self.sheet_vrf = 'vrf'
        self.sheet_subject = 'subject'
        self.sheet_filter = 'filter'
        self.sheet_filter_entr = 'filter_entry'
        self.sheet_vlan_pool = 'vlan_pool'
        self.sheet_vlan_encap_blk = 'vlan_encap_block'
        self.sheet_domain = 'domain'



    def get_excel_file(self):
        self.xl_file = ExcelImporter().import_excelfile()
        return self.xl_file


    def create_bridge_domain(self,xl_file):
        # Bridge SpreadSheet
        excel_file = xl_file
        raw_bridge_domain_sheet = pd.read_excel(excel_file,sheet_name=self.sheet_bridge_domain)
        bridge_domain_sheet = raw_bridge_domain_sheet.sort_values(by=['tenant'])
        raw_bd_subnet_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_bd_subnet)
        bd_subnet_sheet = raw_bd_subnet_sheet.sort_values(by=['tenant'])
        raw_bd_l3out_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_bd_l3out)
        bd_l3out_sheet = raw_bd_l3out_sheet.sort_values(by=['tenant'])
        set_of_tenants = set()
        results = []

        for index, row in bridge_domain_sheet.iterrows():
            if pd.notnull(row['status']):
                continue
            else:
                set_of_tenants.add(row['tenant'])
        set_of_tenants = sorted(set_of_tenants)
        print(set_of_tenants)

        for tenant in set_of_tenants:
            root = Element('polUni')
            # root.set('version', '1.0')
            root.append(Comment('Generated by JTools UCI Builder'))
            self.fvTenant = SubElement(root, 'fvTenant', {'name': tenant, 'status': 'modified'})
            for index, row in bridge_domain_sheet.iterrows():
                if pd.notnull(row['status']):
                    continue
                else:
                    if row['tenant'] == tenant:
                        fvBD = SubElement(self.fvTenant, 'fvBD', {'arpFlood': 'yes','descr': row['description'],
                                                                  'ipLearning':'yes', 'limitIpLearnToSubnets':'yes',
                                                                  'mcastAllow':'no','multiDstPktAct':'bd-flood',
                                                                  'name':row['name'], 'status':'', 'type':'regular',
                                                                  'unicastRoute':'yes', 'unkMacUcastAct':'flood',
                                                                  'unkMcastAct':'flood'})
                        fvRsBDToNdP = SubElement(fvBD,'fvRsBDToNdP',{'tnNdIfPolName':''})
                        fvRsCtx = SubElement(fvBD,'fvRsCtx', {'tnFvCtxName':row['vrf']})
                        fvRsIgmpsn = SubElement(fvBD,'fvRsIgmpsn',{'tnIgmpSnoopPolName':''})
                        fvRsBdToEpRet = SubElement(fvBD,'fvRsBdToEpRet',{'resolveAct':'resolve',
                        'tnFvEpRetPolName':''})
                        for index, row_ip in bd_subnet_sheet.iterrows():
                            if (row_ip['bridge_domain'] == row['name']) and (row_ip['tenant'] == tenant):
                                fvSubnet = SubElement(fvBD, 'fvSubnet', {'ctrl': '','descr':row_ip['description'],
                                                                         'ip':row_ip['bd_subnet'],'preferred':'no',
                                                                         'scope':'public','virtual':'no'})
                        for index, row_l3 in bd_l3out_sheet.iterrows():
                            if (row_l3['bd_name'] == row['name']) and (row_ip['tenant'] == tenant):
                                fvRsBDToOut = SubElement(fvBD, 'fvRsBDToOut', {'tnL3extOutName':row_l3['l3out_name']})
            results.append(root)
        for result in results:
            print(ExcelImporter().prettify(result))
        return results


    def create_contract(self,xl_file):
        # contract SpreadSheet
        excel_file = xl_file
        raw_contract_sheet = pd.read_excel(excel_file,sheet_name=self.sheet_contract)
        contract_sheet = raw_contract_sheet.sort_values(by=['tenant'])
        set_of_tenants = set()
        results = []
        for index, row in contract_sheet.iterrows():
            if pd.notnull(row['status']):
                continue
            else:
                set_of_tenants.add(row['tenant'])
        set_of_tenants = sorted(set_of_tenants)

        for tenant in set_of_tenants:
            root = Element('polUni')
            # root.set('version', '1.0')
            root.append(Comment('Generated by JTools UCI Builder'))
            self.fvTenant = SubElement(root, 'fvTenant', {'name': tenant, 'status': 'modified'})
            for index, row in contract_sheet.iterrows():
                if pd.notnull(row['status']):
                    continue
                else:
                    if row['tenant'] == tenant:
                        vzBrCP = SubElement(self.fvTenant, 'vzBrCP', {'descr': row['description'],
                                                                  'name':row['name'], 'nameAlias':'', 'prio':'unspecified',
                                                                  'scope':row['scope'], 'status':'',
                                                                  'targetDscp':'unspecified'})
            results.append(root)
        for result in results:
            print(ExcelImporter().prettify(result))
        return results


    def create_epg(self,xl_file):
        # EPG SpreadSheet
        excel_file = xl_file
        raw_epg_sheet = pd.read_excel(excel_file,sheet_name=self.sheet_epg)
        epg_sheet = raw_epg_sheet.sort_values(by=['tenant', 'app_profile'])
        epg_subnet_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_epg_subnet)
        raw_epg_domain_ass_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_epg_domain_ass)
        epg_domain_ass_sheet = raw_epg_domain_ass_sheet.sort_values(by=['tenant', 'app_profile'])
        set_of_tenants = set()
        epg_dom_vmm = 'vmm_vmware'
        epg_dom_phy = 'physical'
        epg_dom = ''
        epg_vmm_tDn = ''
        epg_phy_tDN = ''
        results = []

        for index, row in epg_sheet.iterrows():
            if pd.notnull(row['status']):
                continue
            else:
                set_of_tenants.add(row['tenant'])
        set_of_tenants = sorted(set_of_tenants)

        for tenant in set_of_tenants:
            root = Element('polUni')
            # root.set('version', '1.0')
            root.append(Comment('Generated by JTools UCI Builder'))
            self.fvTenant = SubElement(root, 'fvTenant', {'name': tenant, 'status': 'modified'})
            set_of_AppProf = set() 
            for index, row in epg_sheet.iterrows():
                if row['tenant'] == tenant and pd.isnull(row['status']):
                    set_of_AppProf.add(row['app_profile'])
            set_of_AppProf = sorted(set_of_AppProf)
            for app_prof in set_of_AppProf:
                fvAp = SubElement(self.fvTenant,'fvAp',{'name': app_prof,'status':'modified'})
                for index, row_app in epg_sheet.iterrows():
                    if row_app['tenant'] == tenant and row_app['app_profile'] == app_prof:
                        fvAEPg = SubElement(fvAp,'fvAEPg',{'descr': row_app['description'],
                                                                  'floodOnEncap':'disabled', 'isAttrBasedEPg':'no',
                                                                  'matchT':'AtleastOne','name':row_app['name'], 'pcEnfPref':'unenforced', 'prefGrMemb':'exclude',
                                                                  'prio':'unspecified', 'status':''})
                        fvRsBd = SubElement(fvAEPg,'fvRsBd',{'tnFvBDName':row_app['bridge_domain']})
                        for index, row_ip in epg_subnet_sheet.iterrows():
                            if row_ip['epg'] == row_app['name'] and row_ip['tenant'] == row_app['tenant']:
                                fvSubnet = SubElement(fvAEPg,'fvSubnet',{'ctrl':'no-default-gateway','descr':row_ip['description'],
                                                                    'ip':row_ip['epg_subnet'],'preferred':'no','scope':row_ip['subnet_scope'],
                                                                    'virtual':'no'})
                        for index, row_domain in epg_domain_ass_sheet.iterrows():
                            epg_dom = row_domain['domainName']
                            if row_domain['epg_name'] == row_app['name'] and row_domain['tenant'] == row_app['tenant'] and row_domain['domainType'] == epg_dom_vmm:
                                epg_vmm_tDn = 'uni/vmmp-VMware/dom-' + epg_dom
                                fvRsDomAtt_vmm = SubElement(fvAEPg,'fvRsDomAtt',{'encap':'','encapMode':'auto','epgCos':'Cos0','epgCosPref':'disabled',
                                                                    'instrImedcy':'immediate','netflowDir':'both','netflowPref':'disabled',
                                                                    'primaryEncap':'unknown','primaryEncapInner':'unknown','resImedcy':'immediate',
                                                                    'secondaryEncapInner':'unknown','status':'','switchingMode':'native','tDn':epg_vmm_tDn})
                                vmmSecP = SubElement(fvRsDomAtt_vmm,'vmmSecP',{'allowPromiscuous':'reject','descr':'','forgedTransmits':'reject',
                                                                    'macChanges':'reject','name':''})
                            elif row_domain['epg_name'] == row_app['name'] and row_domain['tenant'] == row_app['tenant'] and row_domain['domainType'] == epg_dom_phy:
                                epg_phy_tDN = 'uni/phys-' + epg_dom
                                fvRsDomAtt_phy = SubElement(fvAEPg,'fvRsDomAtt',{'instrImedcy':'immediate','resImedcy':'immediate',
                                                                    'status':'','tDn':epg_phy_tDN})
            results.append(root)
            print(tenant,set_of_AppProf)
            set_of_AppProf.clear()
        for result in results:
            print(ExcelImporter().prettify(result))
        return results


    def create_epg_stat_bnd(self,xl_file):
        # EPG Static Binding SpreadSheet
        excel_file = xl_file
        raw_epg_stat_bnd_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_stat_bnd)
        epg_stat_bnd_sheet = raw_epg_stat_bnd_sheet.sort_values(by=['tenant', 'app_profile'])
        set_of_tenants = set()
        epg_encap = ''
        epg_paths_tDN = ''
        epg_protpaths_tDN = ''
        results = []
        for index, row in epg_stat_bnd_sheet.iterrows():
            if pd.notnull(row['status']):
                continue
            else:
                set_of_tenants.add(row['tenant'])
        set_of_tenants = sorted(set_of_tenants)
        for tenant in set_of_tenants:
            root = Element('polUni')
            root.append(Comment('Generated by JTools UCI Builder'))
            self.fvTenant = SubElement(root,'fvTenant',{'name':tenant,'status':'modified'})
            set_of_AppProf = set() 
            for index, row in epg_stat_bnd_sheet.iterrows():
                if row['tenant'] == tenant and pd.isnull(row['status']):
                    set_of_AppProf.add(row['app_profile'])
            set_of_AppProf = sorted(set_of_AppProf)                   
            for app_prof in set_of_AppProf:
                fvAp = SubElement(self.fvTenant,'fvAp',{'name': app_prof,'status':'modified'})
                for index, row_app in epg_stat_bnd_sheet.iterrows():
                    if row_app['tenant'] == tenant and row_app['app_profile'] == app_prof:
                        fvAEPg = SubElement(fvAp,'fvAEPg',{'matchT':'AtleastOne','name':row_app['name'],'status':'modified'})
                        if pd.isnull(row_app['right_node_id']):    
                            epg_encap = 'vlan-' + str(int(row_app['encap_vlan_id']))
                            epg_paths_tDN = 'topology/pod-' + str(int(row_app['pod_id'])) + '/paths-' + str(int(row_app['left_node_id'])) + '/pathep-[eth' + str(row_app['access_port_id']) + ']'
                            fvRsPathAtt = SubElement(fvAEPg,'fvRsPathAtt',{'descr':'','encap':epg_encap,'instrImedcy':'lazy','mode':row_app['mode'],'status':'','tDn':epg_paths_tDN})
                        else:
                            epg_encap = 'vlan-' + str(row_app['encap_vlan_id'])
                            epg_protpaths_tDN = 'topology/pod-' + str(int(row_app['pod_id'])) + '/protpaths-' + str(int(row_app['left_node_id'])) + '-' + str(int(row_app['right_node_id'])) + '/pathep-[' + row_app['interface_policy_group'] + ']'
                            fvRsPathAtt = SubElement(fvAEPg,'fvRsPathAtt',{'descr':'','encap':epg_encap,'instrImedcy':'lazy','mode':row_app['mode'],'primaryEncap':'unknown','tDn':epg_protpaths_tDN})
            results.append(root)
            print(tenant,set_of_AppProf)
            set_of_AppProf.clear()

        for result in results:
            print(ExcelImporter().prettify(result))
        return results


    def create_app_profile(self,xl_file):
        # App Proile SpreadSheet
        excel_file = xl_file
        raw_app_prof_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_app_prof)
        app_prof_sheet = raw_app_prof_sheet.sort_values(by=['tenant'])
        set_of_tenants = set()
        results = []
        for index, row in app_prof_sheet.iterrows():
            if pd.notnull(row['status']):
                continue
            else:
                set_of_tenants.add(row['tenant'])
        set_of_tenants = sorted(set_of_tenants)

        for tenant in set_of_tenants:
            root = Element('polUni')
            # root.set('version', '1.0')
            root.append(Comment('Generated by JTools UCI Builder'))
            self.fvTenant = SubElement(root, 'fvTenant', {'name': tenant, 'status': 'modified'})
            for index, row in app_prof_sheet.iterrows():
                if pd.notnull(row['status']):
                    continue
                else:
                    if row['tenant'] == tenant:
                        fvAp = SubElement(self.fvTenant, 'fvAp', {'descr':row['description'],'name':row['name'],'prio':'unspecified','status':'',})
                        fvRsApMonPol = SubElement(fvAp,'fvRsApMonPol',{'tnMonEPGPolName':''})
            results.append(root)
        for result in results:
            print(ExcelImporter().prettify(result))
        return results


    def create_tenant(self,xl_file):
        # Tenant SpreadSheet
        excel_file = xl_file
        raw_tenant_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_tenant)
        tenant_sheet = raw_tenant_sheet.sort_values(by=['name'])
        raw_vrf_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_vrf)
        vrf_sheet = raw_vrf_sheet.sort_values(by=['tenant'])
        set_of_tenants = set()

        results = []
        for index, row in tenant_sheet.iterrows():
            if pd.notnull(row['status']):
                continue
            else:
                set_of_tenants.add(row['name'])
        set_of_tenants = sorted(set_of_tenants)

        for tenant in set_of_tenants:
            root = Element('polUni')
            # root.set('version', '1.0')
            root.append(Comment('Generated by JTools UCI Builder'))
            for index, row in tenant_sheet.iterrows():
                if pd.notnull(row['status']):
                    continue
                else:
                    if row['name'] == tenant:
                        self.fvTenant = SubElement(root, 'fvTenant', {'descr':row['description'],'name':row['name'],'status':''})
                        for index, row_vrf in vrf_sheet.iterrows():
                            if row_vrf['tenant'] == tenant and pd.isnull(row['status']):
                                fvCtx = SubElement(self.fvTenant,'fvCtx',{'knwMcastAct':'permit','name':row_vrf['name'],'pcEnfDir':'ingress','pcEnfPref':row_vrf['policy_enforcement'],'status':''})
                                fvRsBgpCtxPol = SubElement(fvCtx,'fvRsBgpCtxPol',{'tnBgpCtxPolName':row_vrf['bgp_timers']})
                                fvRsCtxToExtRouteTagPol = SubElement(fvCtx,'fvRsCtxToExtRouteTagPol',{'tnL3extRouteTagPolName':row_vrf['route_tag_policy']})
                                fvRsOspfCtxPol = SubElement(fvCtx,'fvRsOspfCtxPol',{'tnOspfCtxPolName':row_vrf['ospf_timers']})
                                fvRsCtxToEpRet = SubElement(fvCtx,'fvRsCtxToEpRet',{'tnFvEpRetPolName':''})
                            else:
                                continue
            results.append(root)
        for result in results:
            print(ExcelImporter().prettify(result))
        return results


    def create_subject(self,xl_file):
        # Bridge SpreadSheet
        excel_file = xl_file
        #raw_subject_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_stat_bnd)
        #subject_sheet = raw_subject_sheet.sort_values(by=['tenant'])
        subject_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_subject)
        set_of_tenants = set()
        #sbj_flt_dn = ''
        results = []
        for index, row in subject_sheet.iterrows():
            if pd.notnull(row['status']):
                continue
            else:
                set_of_tenants.add(row['tenant'])
        set_of_tenants = sorted(set_of_tenants)
        #print(set_of_tenants)
        for tenant in set_of_tenants:
            root = Element('polUni')
            root.append(Comment('Generated by JTools UCI Builder'))
            self.fvTenant = SubElement(root,'fvTenant',{'name':tenant,'status':'modified'})
            set_of_contract = set()
            for index, cont_tent in subject_sheet.iterrows():
                if cont_tent['tenant'] == tenant and pd.isnull(cont_tent['status']):
                    set_of_contract.add(cont_tent['contract'])
                else:
                    continue
            set_of_contract = sorted(set_of_contract)
            #print(tenant,set_of_contract)                   
            for contr_sbj in set_of_contract:
                vzBrCP = SubElement(self.fvTenant,'vzBrCP',{'name':contr_sbj,'status':'modified'})
                #set_of_subject = set()
                list_of_subject = []
                #for index, sbjct in subject_sheet.iterrows():
                #    if sbjct['tenant'] == tenant and sbjct['contract'] == contr_sbj and pd.isnull(sbjct['status']):
                #        set_of_subject.add(sbjct['name'])
                #        #if [sbjct['name']] in list_of_subject:
                #        #    continue
                #        #else:
                #        #    append(7)
                #    else:
                #        continue    
                #set_of_subject = sorted(set_of_subject)
                #print('set of contracts',tenant,contr_sbj,set_of_subject)
                #print(len(set_of_subject))
                #create list of subjects to create vzSubj from set_of_subject (unique values only)
                list_of_subject = []
                for index, sbjct2 in subject_sheet.iterrows():
                    line_of_sbj = []
                    if sbjct2['tenant'] == tenant and sbjct2['contract'] == contr_sbj and pd.isnull(sbjct2['status']):
                        line_of_sbj.clear()
                        line_of_sbj = [sbjct2['name'], sbjct2['description'], sbjct2['reverse_filter_port']]
                        if line_of_sbj in list_of_subject:
                            continue
                        else:
                            #print('line of sbj',tenant,contr_sbj,line_of_sbj)
                            list_of_subject.append(line_of_sbj)
                    else:
                        continue
                #for lst_sbj in set_of_subject:
                #    for index, sbjct2 in subject_sheet.iterrows():
                #        line_of_sbj = []
                #        if sbjct2['tenant'] == tenant and sbjct2['contract'] == contr_sbj and sbjct2['name'] == lst_sbj  and pd.isnull(sbjct2['status']):
                #            line_of_sbj.clear()
                #            line_of_sbj = [sbjct2['name'], sbjct2['description'], sbjct2['reverse_filter_port']]
                #            #print('line_of_sbj',tenant,contr_sbj,lst_sbj,line_of_sbj)
                #            list_of_subject.append(line_of_sbj)
                #            break
                list_of_subject.sort()    
                #print('list of subjects',tenant,contr_sbj,list_of_subject)
                for sbj_sbj in list_of_subject:
                    vzSubj = SubElement(vzBrCP,'vzSubj',{'consMatchT':'AtleastOne','descr':sbj_sbj[1],
                                                               'name':sbj_sbj[0],'nameAlias':'','prio':'unspecified',
                                                                'provMatchT':'AtleastOne','revFltPorts':sbj_sbj[2],
                                                                'status':'','targetDscp':'unspecified'})
                    sbj_flt_dn = ''
                    for index, row_filt in subject_sheet.iterrows():
                        if row_filt['tenant'] == tenant and row_filt['contract'] == contr_sbj and row_filt['name'] == sbj_sbj[0] and pd.isnull(row_filt['status']):
                            sbj_flt_dn = 'uni/tn-' + tenant + '/brc-' + contr_sbj + '/subj-' + sbj_sbj[0] + '/rssubjFiltAtt-' + row_filt['filter']
                            #print(sbj_flt_dn)
                            vzRsSubjFiltAtt = SubElement(vzSubj,'vzRsSubjFiltAtt',{'dn':sbj_flt_dn,'tnVzFilterName':row_filt['filter']})
                #set_of_subject.clear()
                list_of_subject.clear()
                            
            results.append(root)
            set_of_contract.clear()
            
        for result in results:
            print(ExcelImporter().prettify(result))
        return results


    def create_filter(self,xl_file):
        # Filter SpreadSheet
        excel_file = xl_file
        #raw_subject_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_stat_bnd)
        #subject_sheet = raw_subject_sheet.sort_values(by=['tenant'])
        raw_flt_sheet = pd.read_excel(excel_file,sheet_name=self.sheet_filter)
        flt_sheet = raw_flt_sheet.sort_values(by=['tenant', 'name'])
        raw_flt_entr_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_filter_entr)
        flt_entr_sheet = raw_flt_entr_sheet.sort_values(by=['tenant', 'filter'])
        set_of_tenants = set()
        flt_fro_port = ''
        flt_to_port = ''
        row_desc = ''
        results = []
        
        for index, row in flt_sheet.iterrows():
            if pd.notnull(row['status']):
                continue
            else:
                set_of_tenants.add(row['tenant'])
        set_of_tenants = sorted(set_of_tenants)
        print(set_of_tenants)

        for tenant in set_of_tenants:
            root = Element('polUni')
            # root.set('version', '1.0')
            root.append(Comment('Generated by JTools UCI Builder'))
            self.fvTenant = SubElement(root, 'fvTenant', {'name': tenant, 'status':'modified'})
            for index, row in flt_sheet.iterrows():
                if pd.notnull(row['status']):
                    continue
                else:
                    if pd.isnull(row['description']):
                        row_desc = ''
                    else:
                        row_desc = str(row['description'])
                    if row['tenant'] == tenant:
                        vzFilter = SubElement(self.fvTenant, 'vzFilter', {'descr':row_desc,'name':row['name'],'status':''})
                        for index, row_flt in flt_entr_sheet.iterrows():
                            if (row_flt['filter'] == row['name']) and (row_flt['tenant'] == tenant) and pd.isnull(row_flt['status']) and pd.notnull(row_flt['from_destination_port']) and pd.notnull(row_flt['to_destination_port']):
                                flt_fro_port = str(int(row_flt['from_destination_port']))
                                flt_to_port = str(int(row_flt['to_destination_port']))
                                vzEntry = SubElement(vzFilter, 'vzEntry', {'dFromPort':flt_fro_port,'dToPort':flt_to_port,
                                                                                'descr':'','etherT':str(row_flt['ether_type']),'name':str(row_flt['name']),'prot':str(row_flt['IP_protocol']),
                                                                                    'stateful':str(row_flt['stateful']),'status':''})
                            else:
                                continue
             
            results.append(root)
        for result in results:
            print(ExcelImporter().prettify(result))
        return results


    def create_vlan_pool(self,xl_file):
        # Bridge SpreadSheet
        excel_file = xl_file
        vlan_pool_sheet = pd.read_excel(excel_file,sheet_name=self.sheet_vlan_pool)
        vlan_encap_blk_sheet = pd.read_excel(excel_file, sheet_name=self.sheet_vlan_encap_blk)
        results = []

        root = Element('polUni')
        # root.set('version', '1.0')
        root.append(Comment('Generated by JTools UCI Builder'))
        self.infraInfra = SubElement(root, 'infraInfra')
        for index, row in vlan_pool_sheet.iterrows():
            if pd.notnull(row['status']):
                continue
            else:
                fvnsVlanInstP = SubElement(self.infraInfra,'fvnsVlanInstP',{'allocMode':row['alloc_mode'],'descr':row['description'],'descr':row['description'],'name':row['name'],'status':''})
                for index, row_vleb in vlan_encap_blk_sheet.iterrows():
                    if pd.isnull(row_vleb['start_vlan']) or pd.isnull(row_vleb['stop_vlan']):
                        continue
                    elif (row_vleb['vlan_pool'] == row['name']) and pd.isnull(row['status']):
                        start_vl = 'vlan-' + str(int(row_vleb['start_vlan']))
                        stop_vl = 'vlan-' + str(int(row_vleb['stop_vlan']))
                        Blk_vl = 'Blk' + str(int(row_vleb['start_vlan'])) + str(int(row_vleb['stop_vlan']))
                        fvnsEncapBlk = SubElement(fvnsVlanInstP, 'fvnsEncapBlk', {'allocMode':row_vleb['alloc_mode'],'from':start_vl,'name':Blk_vl,'role':row_vleb['role'],'to':stop_vl})
        results.append(root)
        for result in results:
            print(ExcelImporter().prettify(result))
        return results

    def create_phy_domain(self,xl_file):
        # Bridge SpreadSheet
        excel_file = xl_file
        domain_sheet = pd.read_excel(excel_file,sheet_name=self.sheet_domain)
        vlan_pool_sheet = pd.read_excel(excel_file,sheet_name=self.sheet_vlan_pool)
        alloc_mode = ''
        phy_dom = ''
        pvl_pool = ''
        results = []

        root = Element('polUni')
        # root.set('version', '1.0')
        root.append(Comment('Generated by JTools UCI Builder'))
        #self.infraInfra = SubElement(root, 'infraInfra')
        for index, row in domain_sheet.iterrows():
            if pd.isnull(row['status']) and (row['type']) == 'physical':
                phy_dom = row['name']
                pvl_pool = row['vlan_pool']
                for index, row_vlp in vlan_pool_sheet.iterrows():
                    if pvl_pool == row_vlp['name']:
                        alloc_mode = str(row_vlp['alloc_mode'])
                        #print('physical domain:',phy_dom,'       ','domain pool:',pvl_pool,'    ','vlan pool',row_vlp['name'])
                        tDNinfraRsVlanNs = 'uni/infra/vlanns-[' + pvl_pool + ']-' + alloc_mode
                        physDomP = SubElement(root,'physDomP',{'name':phy_dom,'status':''})
                        infraRsVlanNs = SubElement(physDomP,'infraRsVlanNs',{'tDn':tDNinfraRsVlanNs})
        results.append(root)
        for result in results:
            print(ExcelImporter().prettify(result))
        return results


    def create_l3_domain(self,xl_file):
        # Bridge SpreadSheet
        excel_file = xl_file
        domain_sheet = pd.read_excel(excel_file,sheet_name=self.sheet_domain)
        vlan_pool_sheet = pd.read_excel(excel_file,sheet_name=self.sheet_vlan_pool)
        alloc_mode = ''
        phy_dom = ''
        pvl_pool = ''
        results = []

        root = Element('polUni')
        # root.set('version', '1.0')
        root.append(Comment('Generated by JTools UCI Builder'))
        #self.infraInfra = SubElement(root, 'infraInfra')
        for index, row in domain_sheet.iterrows():
            if pd.isnull(row['status']) and (row['type']) == 'external_l3':
                l3_dom = row['name']
                pvl_pool = row['vlan_pool']
                for index, row_vlp in vlan_pool_sheet.iterrows():
                    if pvl_pool == row_vlp['name']:
                        alloc_mode = str(row_vlp['alloc_mode'])
                        #print('physical domain:',l3_dom,'       ','domain pool:',pvl_pool,'    ','vlan pool',row_vlp['name'])
                        tDNinfraRsVlanNs = 'uni/infra/vlanns-[' + pvl_pool + ']-' + alloc_mode
                        l3extDomP = SubElement(root,'l3extDomP',{'name':l3_dom,'status':''})
                        infraRsVlanNs = SubElement(l3extDomP,'infraRsVlanNs',{'tDn':tDNinfraRsVlanNs})
        results.append(root)
        for result in results:
            print(ExcelImporter().prettify(result))
        return results




    def generate_xml_files(self):
        pass


    def call_all_aci_elements(self):
        xl_file = Buildetwork().get_excel_file()
        
        #Buildetwork().create_bridge_domain(xl_file)
        #Buildetwork().create_contract(xl_file)
        #Buildetwork().create_epg(xl_file)
        #Buildetwork().create_epg_stat_bnd(xl_file)
        #Buildetwork().create_app_profile(xl_file)
        #Buildetwork().create_tenant(xl_file)
        #Buildetwork().create_subject(xl_file)
        #Buildetwork().create_filter(xl_file)
        #Buildetwork().create_vlan_pool(xl_file)
        #Buildetwork().create_phy_domain(xl_file)
        Buildetwork().create_l3_domain(xl_file)



 






Buildetwork().call_all_aci_elements()





