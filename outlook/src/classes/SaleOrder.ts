class SaleOrder {
    id: number;
    name: string;
    amountTotal: string;
    state: string;
    dateOrder: string;
    requested_at?: string;
    quotation_name?: string;

    static fromJSON(o: Object): SaleOrder {
        const order = new SaleOrder();
        console.log(o, 'sale order info')
        order.id = o['id'];
        order.name = o['name'];
        order.amountTotal = o['amount_total'];
        order.state = o['state'];
        order.dateOrder = o['date_order'];
        order.requested_at = o['requested_at'];
        order.quotation_name = o['quotation_name'];
        return order;
    }
}

export default SaleOrder;
