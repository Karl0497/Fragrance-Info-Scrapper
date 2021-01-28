using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;

namespace FragranceInfo
{
    public class Fragrance
    {
        private IWebDriver Driver;

        private string URL;

        public int Row; //This is for excel mapping

        [Description("Brand")]
        public string Brand;

        [Description("Name")]
        public string Name;

        [Description("Owned")]
        [ReadOnly(true)]
        public string Owned = "No";

        [Description("Sampled")]
        [ReadOnly(true)]
        public string Sampled = "No";

        [Description("Total Votes")]
        public int TotalVotes;

        [Description("Overall Rating")]
        public decimal Rating;

        [Description("Value For Money")]
        public decimal PriceValueRating;

        [Description("Spring")]
        public decimal SpringRating;

        [Description("Summer")]
        public decimal SummerRating;

        [Description("Fall")]
        public decimal FallRating;

        [Description("Winter")]
        public decimal WinterRating;

        [Description("Day")]
        public decimal DayRating;

        [Description("Night")]
        public decimal NightRating;

        [Description("Longevity")]
        public decimal LongevityRating;

        [Description("Sillage")]
        public decimal SillageRating;

        [Description("Gender")]
        public decimal GenderRating;

        public bool IsProcessed;

        public Fragrance(IWebDriver driver, string url)
        {
            Driver = driver;
            URL = url;
        }

        public void GetAllAttributes()
        {

            // Brand
            IWebElement brandElement = Driver.FindElement(By.XPath("//p[@itemprop='brand']//span[@itemprop='name']"));
            Brand = brandElement.Text;
            
            // Name
            IWebElement nameElement = Driver.FindElement(By.XPath("//h1[@itemprop = 'name']"));
            var subElements = nameElement.FindElements(By.XPath(".//*"));
            var text = nameElement.Text;
           
            foreach (IWebElement element in subElements)
            {
                text = text.Replace(element.Text, "");
            }

            Name = text.Trim();

            // Rating
            IWebElement ratingElement = Driver.FindElement(By.XPath("//span[@itemprop='ratingValue']"));
            Rating = Convert.ToDecimal(ratingElement.Text);
            
            // Total Votes
            var ratingCount = Driver.FindElement(By.XPath("//span[@itemprop='ratingCount']"));
            TotalVotes = Convert.ToInt32(ratingCount.Text.Replace(",",""));

            // Ratings based on seasons/day/night 
            WinterRating = GetSeasonalRating("winter");
            SpringRating = GetSeasonalRating("spring");
            SummerRating = GetSeasonalRating("summer");
            FallRating = GetSeasonalRating("fall");
            DayRating = GetSeasonalRating("day");
            NightRating = GetSeasonalRating("night");

            // Attribute ratings
            LongevityRating = GetAttributeRating("LONGEVITY");
            SillageRating = GetAttributeRating("SILLAGE");
            GenderRating = GetAttributeRating("GENDER");
            PriceValueRating = GetAttributeRating("PRICE VALUE");

            IsProcessed = true;
        }

        private decimal GetSeasonalRating(string season)
        {
            IWebElement ratingDiv = Driver.FindElement(By.XPath($"//span[@class='vote-button-legend'][text()='{season}']/../following-sibling::div/div/div"));
            var style = ratingDiv.GetAttribute("style");
            foreach (string property in style.Split(";"))
            {
                var pair = property.Trim().Split(":");

                if (pair[0].Trim() == "width")
                {
                    return Convert.ToDecimal(pair[1].Trim().Replace("%", "")) / 20;
                }
            }

            return 0;
        }

        private decimal GetAttributeRating(string attribute)
        {
            IList<IWebElement> levels = Driver.FindElements(By.XPath($"//span[text() = '{attribute}']/../..//span[@class='vote-button-legend']"));
            int sum = 0;
            int count = 0;

            for (int i = 0; i < levels.Count; i++)
            {
                count += Convert.ToInt32(levels[i].Text);
                sum += Convert.ToInt32(levels[i].Text) * (i + 1);
            }

            return ((decimal)sum / count) / levels.Count * 5; // some attributes have 4 levels only, need to convert to a 5 star system
        }
    }
}
