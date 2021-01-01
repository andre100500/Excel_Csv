using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskAlfa.Models
{
    public class DataModel
    {
        public class Currency
        {
            public Dictionary<string, double> currency = new Dictionary<string, double>();

            public Currency addValueToCurrency(string currencyName, double value)
            {
                if (currency.ContainsKey(currencyName))
                {
                    double oldVal = currency[currencyName];
                    currency[currencyName] = oldVal + value;
                } else
                {
                    currency.Add(currencyName, value);
                }
                return this;
            }

            public double getValueByCurrencyName(string currencyName)
            {
                if (currency.ContainsKey(currencyName))
                {
                    return currency[currencyName];
                }
                else
                {
                    return 0;
                }
            }
        }

        public Dictionary<string, Currency> accounts = new Dictionary<string, Currency>();

        public double getValue(string account, string currency)
        {
            if (accounts.ContainsKey(account))
            {
                return accounts[account].getValueByCurrencyName(currency);
            }
            else
            {
                return 0;
            }
        }

        public void setValue(string account, string currency, double value)
        {
            if (accounts.ContainsKey(account))
            {
                accounts[account].addValueToCurrency(currency, value);
            } else
            {
                accounts.Add(account, new Currency().addValueToCurrency(currency, value));
            }
        }
    }
}
