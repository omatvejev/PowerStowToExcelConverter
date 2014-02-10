using System;
using System.Collections;

namespace PowerStowToExcelConverter.Model
{
    class ShipmentCompany
    {
        private string name;
        private ArrayList shipments;

        public ShipmentCompany(string name)
        {
            this.name = name;
            shipments = new ArrayList();
        }

        public string Name
        {
            get 
            { 
                return this.name; 
            }
        }

        public ArrayList Shipments
        {
            get
            {
                return this.shipments;
            }
        }

        public void addShipment(string location)
        {
            ShipmentData shipment = new ShipmentData(location);
            this.shipments.Add(shipment);
        }
    }
}
