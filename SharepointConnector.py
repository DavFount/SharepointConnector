import requests

class SharepointConnector:
    access_token = None
    base_url = None

    def __init__(self, client_id, client_secret, tenet_id, resource_id, base_url):
        url = f'https://accounts.accesscontrol.windows.net/{tenet_id}/tokens/OAuth/2'
        x = requests.post(url, data={
            'grant_type': 'client_credentials',
            'client_id': f'{client_id}@{tenet_id}',
            'client_secret': f'{client_secret}',
            'resource': f'{resource_id}/sacent0.sharepoint.com@{tenet_id}'

        }, headers= {
            'Content-Type': 'application/x-www-form-urlencoded'
        })

        response = eval(x.text)
        self.access_token = response.get('access_token')
        self.base_url = base_url

    def GetListItems(self, site, list_name, top=None):
        if(top):
            url = f"https://{self.base_url}/sites/{site}/_api/lists/GetByTitle('{list_name}')/items?$top={top}"
        else:
            url = f"https://{self.base_url}/sites/{site}/_api/lists/GetByTitle('{list_name}')/items"

        response = requests.get(url, headers={
            'Authorization': 'Bearer ' + self.access_token,
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/x-www-form-urlencoded'
        })
        return response

    def GetListItemsByColumn(self, site, list_name, column_name, value):
        url = f"https://{self.base_url}/sites/{site}/_api/lists/GetByTitle('{list_name}')/items?$filter={column_name} eq '{value}'"
        response = requests.get(url, headers={
            'Authorization': 'Bearer ' + self.access_token,
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/x-www-form-urlencoded'
        })
        return response

    def UpdateListItemById(self, site, list_name, entry_id, item_object):
        url = f"https://{self.base_url}/sites/{site}/_api/lists/GetByTitle('{list_name}')/Items({entry_id})"
        response = requests.post(url, data = item_object, headers= {
            'Authorization': 'Bearer ' + self.access_token,
            'X-RequestDigest': 'form digest value',
            'If-Match': '*',
            'X-HTTP-Method': 'MERGE',
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
        })
        return response
    
    def InsertListItem(self, site, list_name, item_object):
        url = f"https://{self.base_url}/sites/{site}/_api/lists/GetByTitle('{list_name}')/items"
        response = requests.post(url, data = item_object, headers= {
            'Authorization': 'Bearer ' + self.access_token,
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': 'form digest value'
        })
        return response
