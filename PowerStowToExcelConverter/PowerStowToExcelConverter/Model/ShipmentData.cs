using System;
using System.Collections.Generic;

namespace PowerStowToExcelConverter.Model
{
    class ShipmentData
    {
        // Creating the containers structure to stores the number of various containers are being shipped
        public struct Containers
        {
            public int twenty;
            public int forty;
            public int fortyHC;
            public int fortyFive;
            public double totalWeight;
        };

        private string location;
        private Containers fullContainers;
        private Containers emptyContainers;

        public ShipmentData(string location)
        {
            this.location = location;
        }

        public string Location
        {
            get
            {
                return this.location;
            }
        }

        public Containers FullContainers
        {
            get
            {
                return this.fullContainers;
            }
        }

        public Containers EmptyContainers
        {
            get
            {
                return this.emptyContainers;
            }
        }

        public void createFullContainers(string input)
        {
            string[] containers = input.Split('/');
            try
            {
                fullContainers.twenty = int.Parse(containers[0]);
                fullContainers.forty = int.Parse(containers[1]);
                fullContainers.fortyHC = int.Parse(containers[2]);
                fullContainers.fortyFive = int.Parse(containers[3]);
                fullContainers.totalWeight = double.Parse(containers[4]);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void createEmptyContainers(string input)
        {
            string[] containers = input.Split('/');
            try
            {
                emptyContainers.twenty = int.Parse(containers[0]);
                emptyContainers.forty = int.Parse(containers[1]);
                emptyContainers.fortyHC = int.Parse(containers[2]);
                emptyContainers.fortyFive = int.Parse(containers[3]);
                emptyContainers.totalWeight = double.Parse(containers[4]);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
