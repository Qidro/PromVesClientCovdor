using System;
using System.Collections.Generic;
using System.Text;

namespace PromVesClient.DTO
{
    public class WeighingDto
    {
        
        public decimal Platform1Left { get; set; }

        public decimal Platform1Right { get; set; } 
        public decimal Platform2Left { get; set; }
        public decimal Platform2Right { get; set; }
        public string VagonNumber { get; set; }
        public decimal TareWeight { get; set; }
        public decimal GrossWeight { get; set; }

        public string TypeWeighing { get; set; }
        //грузоотправитель
        public string? Shipper { get; set; }
        //грузополучатель
        public string? Consignee { get; set; }
        //груз
        public string? Cargo { get; set; }
        //номер накладной
        public string? InvoiceNumber { get; set; }
        //номер накладной
        public DateTime? InvoiceDateTime { get; set; }
        //вес по накладной
        public decimal? InvoiceWeighing { get; set; }
        public Guid IdReceipt { get; set; }
    }
}
