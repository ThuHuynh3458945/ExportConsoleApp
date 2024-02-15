namespace NFQ_Kata
{
    public class Item
    {
        public string Name { get; set; }
        public int SellIn { get; set; }
        public int Quality { get; set; }
        public override string ToString()
        {
            return $"{Name}, {SellIn}, {Quality}";
        }
    }
    public class GildedRose
    {
        private IList<Item> Items;
        public GildedRose(IList<Item> Items)
        {
            this.Items = Items;
        }
        public void UpdateQuality()
        {
            foreach (var item in Items)
            {
                // "Sulfuras", being a legendary item, never has to be sold or decreases in Quality
                if (item.Name != "Sulfuras, Hand of Ragnaros")
                {
                    item.SellIn--;

                    if (item.Name == "Aged Brie")
                    {
                        item.Quality++;
                    }
                    else if (item.Name == "Backstage passes to a TAFKAL80ETC concert")
                    {
                        if (item.SellIn <= 0)
                        {
                            item.Quality = 0;
                        }
                        // Quality increases by 3 when there are 5 days or less but Quality drops to 0 after the concert
                        else if (item.SellIn <= 5)
                        {
                            item.Quality += 3;
                        }
                        // Quality increases by 2 when there are 10 days or less
                        else if (item.SellIn <= 10)
                        {
                            item.Quality += 2;
                        }
                        else
                        {
                            item.Quality++;
                        }
                    }
                    //"Conjured" items degrade in Quality twice as fast as normal items Feel free to make any changes to the 
                    else if (item.Name == "Conjured Mana Cake")
                    {
                        item.Quality -= 2;
                    }
                    else 
                    {
                        item.Quality--;
                    }

                    // Ensure quality is within bounds
                    // The Quality of an item is never more than 50
                    item.Quality = Math.Max(0, Math.Min(50, item.Quality));

                    if (item.SellIn < 0)
                    {
                        // "Aged Brie" actually increases in Quality the older it gets
                        if (item.Name == "Aged Brie")
                        {
                            item.Quality++;
                        }
                        // less but Quality drops to 0 after the concert
                        else if (item.Name == "Backstage passes to a TAFKAL80ETC concert")
                        {
                            item.Quality = 0;
                        }
                        else if (item.Name != "Sulfuras, Hand of Ragnaros")
                        {
                            //"Conjured" items degrade in Quality twice as fast as normal items Feel free to make any changes to the 
                            if (item.Name == "Conjured Mana Cake")
                            {
                                item.Quality -= 2;
                            }
                            else
                            {
                                item.Quality--;
                            }
                        }
                        // Ensure quality is within bounds after the sell-by date
                        // The Quality of an item is never more than 50
                        item.Quality = Math.Max(0, Math.Min(50, item.Quality));
                    }
                }
            }
        }
    }
}
