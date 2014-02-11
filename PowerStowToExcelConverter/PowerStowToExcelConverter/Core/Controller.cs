using System;
using System.Collections.Generic;
using PowerStowToExcelConverter.Model;
using System.Collections;

namespace PowerStowToExcelConverter.Core
{
    // Singleton class that controls all the actions that occur in the program
    class Controller : Exception
    {
        private static Controller instance;
        private Terminal terminal;
        private ShipmentCompany[] companies;
        private Controller() { }
        private Translator translator;

        public static Controller Instance
        {
            get
            {
                // Create a new instance if it doesn't exist
                if (instance == null)
                {
                    instance = new Controller();
                }
                return instance;
            }
        }

        public Translator Translator
        {
            get
            {
                return translator;
            }
        }
        public void loadTranslator()
        {
            // Create a translator object. In-case if there is any problem then set the object to null
            try
            {
                translator = new Translator(@"Translation.xml");
            }
            catch (Exception ex)
            {
                translator = null;
                throw ex;
            }
        }

        public void readFile(string filename, string additionalOptions)
        {
            Reader reader = new Reader(filename);
            createTerminal(reader);
            createShipmentCompanies(reader, additionalOptions);
            createShipmentLocations(reader);
            createContainers(reader);
        }

        private void createTerminal(Reader reader)
        {
            string terminalName = reader.parseTerminalInformation();
            string[] vesselAndVoyage = reader.parseVesselandVoyageInformation();

            // Could not find the terminal name there might a problem with the file
            if (terminalName == null || vesselAndVoyage == null)
            {
                throw new DataMisalignedException("Could not parse the file!");
            }

            // Create the new terminal
            terminal = new Terminal(terminalName, vesselAndVoyage[0], vesselAndVoyage[1]);
        }

        private void createShipmentCompanies(Reader reader, string additionalOptions)
        {
            string[] companyNames = reader.parseCompaniesInformation();

            ArrayList companyList = new ArrayList();

            // Create the new companies
            foreach (string name in companyNames)
            {
                companyList.Add(new ShipmentCompany(name));
            }

            // Handle the additional options
            if (!additionalOptions.Equals(""))
            {
                string[] temp = additionalOptions.Split(',');
                foreach (string name in temp)
                {
                    // Try to remove any junk input that the user might try to input
                    if (!name.Equals(',') && !name.Equals(""))
                        companyList.Add(new ShipmentCompany(name.Trim()));
                }
            }

            // Convert the array list into an array
            this.companies = companyList.ToArray(typeof(ShipmentCompany)) as ShipmentCompany[];
        }

        private void createShipmentLocations(Reader reader)
        {
            string[] portNames = reader.parseShipmentPorts();

            // Add ports to each company using deep copy
            foreach (ShipmentCompany company in this.companies)
            {
                foreach (String port in portNames)
                {
                    // create a new shipment. 
                    // Note: The full and empty containers are automaticly assigned a default value
                    // due to the C# global variable structure.
                    company.addShipment(port);
                }
            }
        }

        private void createContainers(Reader reader)
        {
            string[][] fullContainersData;
            string[][] emptyContainersData;

            foreach (ShipmentCompany company in companies)
            {
                fullContainersData = reader.parseFullContainers(company.Name);

                int i = 0;

                // Loop through each container data and determine which ports need to have their default value changed
                while (fullContainersData != null && i < fullContainersData[0].Length)
                {
                    foreach (ShipmentData shipment in company.Shipments)
                    {
                        // This shipment location has a container data
                        if (shipment.Location.Equals(fullContainersData[0][i]))
                        {
                            shipment.createFullContainers(fullContainersData[1][i]);
                            break;
                        }
                    }
                    i++;
                }           

                emptyContainersData = reader.parseEmptyContainers(company.Name);

                i = 0; // Reset the iterator

                // Loop through each container data and determine which ports need to have their default value changed
                while (emptyContainersData != null && i < emptyContainersData[0].Length)
                {
                    foreach (ShipmentData shipment in company.Shipments)
                    {
                        // This shipment location has a container data
                        if (shipment.Location.Equals(emptyContainersData[0][i]))
                        {
                            shipment.createEmptyContainers(emptyContainersData[1][i]);
                            break;
                        }
                    }
                    i++;
                }            
            }
        }


        public void writeFile(string path)
        {
            Writer writer = new Writer(path, true);

            writer.writeTerminalInformation(terminal, 1, 1);

            // Shipment writer settings
            int i = 0;
            int row = 3;
            int col = 1;

            // Send the shipment information to the writer
            foreach (ShipmentCompany company in companies)
            {
                // Only write the heading information in the starting row
                if (row == 3) 
                    writer.writeShipmentInformation(company, row, col, true);
                else
                    writer.writeShipmentInformation(company, row, col, false);

                i++;

                // Check if a new row is required for the next output
                if (i % 2 == 0)
                {
                    row = row + company.Shipments.Count + 2;
                    col = 1;
                }
                else
                    col = 21;
            }
            writer.save();
        }
    }
}
