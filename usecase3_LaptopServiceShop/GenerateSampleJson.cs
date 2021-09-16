using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

public class GenerateSampleJson
{
    public static void doIt()
    {
        // add two software service orders into software order array
        var OrderList = new ArrayList();
        var order = new ServiceOrder();
        order.order_id = "S000-001";
        order.laptop_model = "Dell XPS 15";
        order.service_type = "software";
        order.problem_description = "running too slow";
        order.customer_name = "Lisha Li";
        order.email = "ls@hotmail.com";
        order.cost = 0;
        order.invoice_number = 0;
        OrderList.Add (order);

        order = new ServiceOrder();
        order.order_id = "S000-002";
        order.laptop_model = "Thinkpad X1";
        order.service_type = "software";
        order.problem_description = "Recover Data in C drive";
        order.customer_name = "Michael Wang";
        order.email = "mw@gmail.com";
        order.cost = 0;
        order.invoice_number = 0;
        OrderList.Add (order);

        order = new ServiceOrder();
        order.order_id = "H000-101";
        order.laptop_model = "MacBook Air Gold";
        order.service_type = "hardware";
        order.problem_description = "broken keyboard";
        order.customer_name = "Robert Wang";
        order.email = "rbt@yahoo.com";
        order.cost = 0;
        order.invoice_number = 0;
        OrderList.Add (order);

        order = new ServiceOrder();
        order.order_id = "H000-102";
        order.laptop_model = "Lenova Yoga 7i";
        order.service_type = "hardware";
        order.problem_description = "extend memory";
        order.customer_name = "Bruce Li";
        order.email = "blue@facebook.com";
        order.cost = 0;
        order.invoice_number = 0;
        OrderList.Add (order);

        // save order list to json file.
        String jsonString = JsonSerializer.Serialize(OrderList);
        File
            .WriteAllText(Path
                .Combine("Resources", "DataSource", "serviceOrders.json"),
            jsonString);
    }
}
