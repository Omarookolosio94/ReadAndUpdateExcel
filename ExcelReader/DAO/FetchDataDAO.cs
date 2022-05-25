using ExcelReader.Models;
using Newtonsoft.Json;
using System.Net.Http;

namespace ExcelReader.DAO
{
    public class FetchDataDAO
    {
        readonly string baseurl = "https://basket-api.azurewebsites.net/api/v1/";
        public Product GetProduct(int productId)
        {
            Product product = new();

            try
            {
                using var client = new HttpClient();

                client.BaseAddress = new Uri(baseurl);

                var getTask = client.GetAsync($"products?productID={productId}");

                getTask.Wait();

                var result = getTask.Result;

                if (result.IsSuccessStatusCode)
                {
                    var readTask = result.Content.ReadAsStringAsync();

                    string postStatus = result.IsSuccessStatusCode.ToString();

                    if (readTask.Result != null)
                    {
                        product = JsonConvert.DeserializeObject<BaseResponse<Product>>(readTask.Result).Data;
                    }
                }
                else
                {
                    product = null;
                }
            }
            catch (Exception ex)
            {
                product = null;
            }

            return product;
        }

        public Ratings GetRatings(int productId)
        {
            Ratings ratings = new();

            try
            {
                using var client = new HttpClient();

                client.BaseAddress = new Uri(baseurl);

                var getTask = client.GetAsync($"products/rating?productID={productId}");

                getTask.Wait();

                var result = getTask.Result;

                if (result.IsSuccessStatusCode)
                {
                    var readTask = result.Content.ReadAsStringAsync();

                    string postStatus = result.IsSuccessStatusCode.ToString();

                    if (readTask.Result != null)
                    {
                        ratings = JsonConvert.DeserializeObject<BaseResponse<Ratings>>(readTask.Result).Data;
                    }
                }
                else
                {
                    ratings = null;
                }
            }
            catch (Exception ex)
            {
                ratings = null;
            }

            return ratings;
        }

    }
}
