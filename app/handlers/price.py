import requests

class PriceHandler:
    def __init__(self, base_url):
        self.base_url = base_url

    def search(self, query):
        """Search for products based on a query string."""
        response = requests.get(self.base_url, params={'search': query})
        
        if response.status_code == 200:
            return response.json()  # Return the JSON response if successful
        else:
            response.raise_for_status()  # Raise an error for bad responses

if __name__ == "__main__":
    # Example usage
    handler = PriceHandler("https://b2b.resurs-media.ru/netshop/")
    try:
        result = handler.search("example product")
        print(result)
    except Exception as e:
        print(f"An error occurred: {e}")
