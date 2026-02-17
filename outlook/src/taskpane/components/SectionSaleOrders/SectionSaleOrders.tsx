import * as React from 'react';
import Partner from '../../../classes/Partner';
import AppContext from '../AppContext';
import SaleOrder from '../../../classes/SaleOrder';
import Section from '../Section/Section';
import { _t } from '../../../utils/Translator';
import './SectionSaleOrders.css';

type SaleOrderSectionProps = {
    partner: Partner;
    canCreatePartner: boolean;
    onSeeAll: () => void;
};

class SectionSaleOrders extends React.Component<SaleOrderSectionProps> {
    private getOrderDescription = (order: SaleOrder): string => {
        return order.amountTotal + (order.quotation_name ? ' - ' + order.quotation_name : '');
    };

    private getTitle = (order: SaleOrder): string => {
        return order.name + ' (' + order.state + ') ';
    };

    render() {
        const allSaleOrders = this.props.partner.saleOrders || [];
        const totalCount = allSaleOrders.length;

        let displayedOrders = [];
        if (totalCount > 0) {
            const current = allSaleOrders.find(o => o.id === this.props.partner.currentSaleOrderId);
            displayedOrders = [current || allSaleOrders[0]];
        }

        displayedOrders = displayedOrders.map((order) => {
            return {
                ...order,
                name: this.getTitle(order),
            };
        });

        return (
            <Section
                records={displayedOrders}
                recordCount={totalCount}
                partner={this.props.partner}
                canCreatePartner={this.props.canCreatePartner}
                model="sale.order"
                odooEndpointCreateRecord={null}
                odooRecordIdName="id"
                odooRedirectAction="sale.action_orders"
                title="Quotations"
                titleCount="Quotations (%(count)s)"
                msgNoPartner="Save Contact to see Quotations."
                msgNoPartnerNoAccess="Access denied to quotations."
                msgNoRecord="No quotations found for this contact."
                msgLogEmail="Link to Odoo"
                getRecordDescription={this.getOrderDescription}
                getRecordHasValue={(order) => !!order.requested_at}>
                {allSaleOrders.length > 0 && (
                    <div className="see-all-sale-orders" onClick={this.props.onSeeAll}>
                        {_t('See All')}
                    </div>
                )}
            </Section>
        );
    }
}

SectionSaleOrders.contextType = AppContext;

export default SectionSaleOrders;
