import requests, ipaddress
from pathlib import Path
from microsoft_pac.exceptions import *
"""
see
https://docs.microsoft.com/en-us/archive/blogs/onthewire/new-office-365-url-categories-to-help-you-optimize-the-traffic-which-really-matters
"""
directProxyVarName = "direct"
defaultProxyVarName = "proxyServer"
client_request_id = "b10c5ed1-bad1-445f-b386-b919946339a7"

class MicrosoftPac(object):
    def __init__(self, instance = 'Worldwide', required=True, tenant=None, category=0, verbose=False):
        self.verbose = verbose
        self.tenant = tenant
        if self.verbose:
            print("VERBOSE:\tRetrieving Microsoft Endpoint information")
        self.endpoints = self.get_services(instance)
        self.required = required
        self.category = category
        self.category_set = {
            0: ['Default', 'Allow', 'Optimize'],
            1: ['Allow'],
            2: ['Allow', 'Optimize']
        }

        self.extract()

    def get_services(self, instance = 'Worldwide'):
        instance_set = ['Worldwide', 'Germany', 'China', 'USGovDoD', 'USGovGCCHigh']
        if instance not in instance_set:
            raise UnknownValue(instance, instance_set)

        microsoft_url = f"https://endpoints.office.com/endpoints/{instance}"
        params = {
                "ClientRequestId": client_request_id
            }
        if self.tenant:
            params.update({'TenantName': self.tenant.lower()})

        try:
            if self.verbose:
                pritn(f"VERBOSE:\tGET {microsoft_url} with {params}")
            r = requests.get(microsoft_url, params=params, verify=True)
        except requests.exceptions.RequestException as err:
            raise SystemExit(err)
        
        return r.json()

    def extract(self):
        service_set = ['Exchange', 'Skype', 'SharePoint', 'Common']
        categories = ['Default', 'Allow', 'Optimize']
        webtraffic = ['80', '80,443', '443']
        # figure out how to split out udp ports special tcp and 80/443
        service_dict = {key: {c: {'urls': [], 'ips': []} for c in categories} for key in service_set}
        #service_dict = {key: {c: {'urls': [], 'ips': {}} for c in categories} for key in service_set}


        if self.required == False:
            for i in self.endpoints:
                if i.get('urls') != None:
                    for url in i.get('urls'):
                        service_dict[i.get('serviceArea')].append(url)    
        else:
            for i in self.endpoints:
                if i.get('required') == self.required:
                    #if i.get('category', 'Default') in self.category_set[self.category]: 
                    #    if i.get('urls') != None:
                    #        for url in i.get('urls'):
                    #            service_dict[i.get('serviceArea')].append(url) 
                    if i.get('urls') != None:
                        for url in i.get('urls'):
                            service_dict[i.get('serviceArea')][i.get('category', 'Default')]['urls'].append(url)
                    if i.get('ips') != None:
                        if i.get('tcpPorts', 'None') in webtraffic:
                            for ips in i.get('ips'):
                                service_dict[i.get('serviceArea')][i.get('category', 'Default')]['ips'].append(ips)
                                #service_dict[i.get('serviceArea')][i.get('category', 'Default')]['ips'].update({})
        
        #if self.tenant:
        #    # Movies *.sharepoint.com into 'Default' placing tenant traffic in 'Optimize'
        #    # This will allow no decryption on the Optimize traffic, but still decrypt any
        #    # sharepoint that will share with other devices
        #    #service_dict['SharePoint']['Optimize']['urls'].append(f"{self.tenant.lower()}.sharepoint.com")
        #    #service_dict['SharePoint']['Optimize']['urls'].append(f"{self.tenant.lower()}-my.sharepoint.com")
        #    for share in service_dict['SharePoint']['Optimize']['urls']:
        #        service_dict['SharePoint']['Default']['urls'].append(share)
        #
        #    service_dict['SharePoint']['Optimize'].update({'urls': [
        #        f"{self.tenant.lower()}-my.sharepoint.com",
        #        f"{self.tenant.lower()}.sharepoint.com"
        #    ]})


        self.endpoints_breakdown = service_dict
    
    def create_firewall_acl(self, location):
        """
        Creates a Palo Alto format ACL csv to upload into pa sync script\n
        IP_Address,Name,Description
        """
        service_set = ['Exchange', 'Skype', 'SharePoint', 'Common']
        allow_list = {key: [] for key in service_set}
        optimize_list = {key: [] for key in service_set}
        default_list = {key: [] for key in service_set}
        list_of_lists = {'Allow': allow_list, 'Optimize': optimize_list, 'Default': default_list}

        for service in service_set:
            for category, listname in list_of_lists.items():
                for ip_addr in self.endpoints_breakdown[service][category]['ips']:
                        list_of_lists[category][service].append([ip_addr, f"Azure_{service}_{category}_", "Description here"])

        self.allow_list = allow_list
        self.optimize_list = optimize_list
        self.default_list = default_list


    def create_pac_file(self):
        pac_file = [
            "// This PAC file will provide proxy config to Microsoft 365 services",
            "//  using data from the public web service for all endpoints",
            "function FindProxyForURL(url, host)",
            "{",
        ]
        service_set = ['Exchange', 'Skype', 'SharePoint', 'Common']
        categories = ['Default', 'Allow', 'Optimize']

        tmp_array = []
        for service in service_set:
            for category in categories:
                for url in self.endpoints_breakdown[service][category]['urls']:
                    tmp_array.append(f"host, \"{url}\"")

        for i in range(len(tmp_array)):
            if i == 0:
                pac_file.append(f"\tif(shExpMatch({tmp_array[i]})")
            else:
                pac_file.append(f"\t\t|| shExpMatch({tmp_array[i]})")
        pac_file.append("\t\t)")
        pac_file.append("\t{")
        pac_file.append("\t\treturn DIRECT;")
        pac_file.append("\t}")
        pac_file.append("}")

        self.pac_file = pac_file

def get_services(self, instance = 'Worldwide'):
    instance_set = ['Worldwide', 'Germany', 'China', 'USGovDoD', 'USGovGCCHigh']
    if instance not in instance_set:
        raise UnknownValue(instance, instance_set)

    microsoft_url = f"https://endpoints.office.com/endpoints/{instance}"
    params = {
            "ClientRequestId": client_request_id
        }

    try:
        r = requests.get(microsoft_url, params=params, verify=True)
    except requests.exceptions.RequestException as err:
        raise SystemExit(err)
    
    return r.json()
