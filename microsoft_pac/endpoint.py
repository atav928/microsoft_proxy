import requests, ipaddress
from microsoft_pac.exceptions import *
"""
see
https://docs.microsoft.com/en-us/archive/blogs/onthewire/new-office-365-url-categories-to-help-you-optimize-the-traffic-which-really-matters
"""
directProxyVarName = "direct"
defaultProxyVarName = "proxyServer"
client_request_id = "b10c5ed1-bad1-445f-b386-b919946339a7"

class MicrosoftPac(object):
    def __init__(self, instance = 'Worldwide', required=True, tenant=None, category=0):
        self.endpoints = self.get_services(instance)
        self.required = required
        self.tenant = tenant
        self.category = category
        self.category_set = {
            0: ['Default', 'Allow', 'Optimize'],
            1: ['Allow'],
            2: ['Allow', 'Optimize']
        }

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

    def extract(self):
        service_set = ['Exchange', 'Skype', 'SharePoint', 'Common']
        categories = ['Default', 'Allow', 'Optimize']
        service_dict = {key: {c: {'urls': [], 'ips': []} for c in categories} for key in service_set}

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
                        for ips in i.get('ips'):
                            service_dict[i.get('serviceArea')][i.get('category', 'Default')]['ips'].append(ips)
        
        if self.tenant:
            # Movies *.sharepoint.com into 'Default' placing tenant traffic in 'Optimize'
            # This will allow no decryption on the Optimize traffic, but still decrypt any
            # sharepoint that will share with other devices
            #service_dict['SharePoint']['Optimize']['urls'].append(f"{self.tenant.lower()}.sharepoint.com")
            #service_dict['SharePoint']['Optimize']['urls'].append(f"{self.tenant.lower()}-my.sharepoint.com")
            for share in service_dict['SharePoint']['Optimize']['urls']:
                service_dict['SharePoint']['Default']['urls'].append(share)

            service_dict['SharePoint']['Optimize'].update({'urls': [
                f"{self.tenant.lower()}-my.sharepoint.com",
                f"{self.tenant.lower()}.sharepoint.com"
            ]})


        self.endpoints_breakdown = service_dict

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