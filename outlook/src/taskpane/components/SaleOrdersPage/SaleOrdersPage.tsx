import * as React from 'react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faScroll } from '@fortawesome/free-solid-svg-icons';
import Partner from '../../../classes/Partner';
import SaleOrder from '../../../classes/SaleOrder';
import ListItem from '../ListItem/ListItem';
import { _t } from '../../../utils/Translator';
import './SaleOrdersPage.css';
import AppContext from '../AppContext';
import { ContentType, HttpVerb, sendHttpRequest } from '../../../utils/httpRequest';
import api from '../../api';
import Lead from '../../../classes/Lead';
import Task from '../../../classes/Task';
import HelpdeskTicket from '../../../classes/HelpdeskTicket';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { OdooTheme } from '../../../utils/Themes';

type SaleOrdersPageProps = {
    partner: Partner;
    onBack: () => void;
    loadPartner: boolean;
    onPartnerChanged?: (Partner) => void;
};

type SaleOrdersPageState = {
    partner: Partner;
    isLoading: boolean;
};

class SaleOrdersPage extends React.Component<SaleOrdersPageProps, SaleOrdersPageState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            partner: props.partner,
            isLoading: props.loadPartner,
        };
    }

    private fetchContact = () => {
        const partner = this.state.partner;

        const requestData = partner.isAddedToDatabase()
            ? { partner_id: partner.id }
            : { email: partner.email, name: partner.name };

        const partnerRequest = sendHttpRequest(
            HttpVerb.POST,
            api.baseURL + api.getPartner,
            ContentType.Json,
            this.context.getConnectionToken(),
            requestData,
            true,
        );

        this.context.addRequestCanceller(partnerRequest.cancel);

        partnerRequest.promise
            .then((response) => {
                const parsed = JSON.parse(response);

                const newPartner =
                    parsed.result && parsed.result.partner && Object.keys(parsed.result.partner).length
                        ? Partner.fromJSON(parsed.result.partner)
                        : Partner.fromJSON({ email: partner.email, name: partner.name });

                if (parsed.result.leads) {
                    newPartner.leads = parsed.result.leads.map((lead_json) => Lead.fromJSON(lead_json));
                }
                if (parsed.result.tasks) {
                    newPartner.tasks = parsed.result.tasks.map((task_json) => Task.fromJSON(task_json));
                }
                if (parsed.result.tickets) {
                    newPartner.tickets = parsed.result.tickets.map((ticket_json) =>
                        HelpdeskTicket.fromJSON(ticket_json),
                    );
                }
                if (parsed.result.sale_orders) {
                    newPartner.saleOrders = parsed.result.sale_orders
                        .map((so_json) => SaleOrder.fromJSON(so_json));
                }
                if (parsed.result.user_companies) {
                    this.context.setUserCompanies(parsed.result.user_companies);
                }

                const canCreatePartner = parsed.result.can_create_partner !== false;
                this.context.setCanCreatePartner(canCreatePartner);

                this.setState({
                    partner: newPartner,
                    isLoading: false,
                });
                if (this.props.onPartnerChanged) {
                    this.props.onPartnerChanged(newPartner);
                }
            })
            .catch((error) => {
                this.context.showHttpErrorMessage(error);
                this.setState({ isLoading: false });
            });
    };

    componentDidMount() {
        if (this.props.loadPartner && this.context.isConnected()) {
            this.fetchContact();
        }
    }

    private getOrderDescription = (order: SaleOrder): string => {
        return order.amountTotal + (order.quotation_name ? ' - ' + order.quotation_name : '');
    };

    private getTitle = (order: SaleOrder): string => {
        return order.name + ' (' + order.state + ') ';
    };

    render() {
        const { partner, isLoading } = this.state;
        const saleOrders = partner.saleOrders || [];

        if (isLoading) {
            return <Spinner className="contact-spinner" size={SpinnerSize.large} theme={OdooTheme} />;
        }

        return (
            <div className="sale-orders-page">
                <div className="sale-orders-header">
                    <FontAwesomeIcon icon={faScroll} className="header-icon" />
                    <h2>{_t('Quotations (%(count)s)', { count: saleOrders.length.toString() })}</h2>
                </div>

                <div className="sale-orders-list">
                    {saleOrders.length > 0 ? (
                        <>
                            {saleOrders.filter(o => o.id === partner.currentSaleOrderId).map((order) => (
                                <div key={`current-${order.id}`}>
                                    <div className="recent">{_t('Recent')}</div>
                                    <ListItem
                                        model="sale.order"
                                        res_id={order.id}
                                        title={this.getTitle(order)}
                                        description={this.getOrderDescription(order)}
                                        logTitle={_t('Link to Odoo')}
                                        partnerId={partner.id}
                                        hasValue={!!order.requested_date}
                                    />
                                </div>
                            ))}

                            {(saleOrders.filter(o => o.id !== partner.currentSaleOrderId).length > 0) && (
                                <div className="sale-order-section-label other-label">{_t('Others')}</div>
                            )}

                            {saleOrders.filter(o => o.id !== partner.currentSaleOrderId).map((order) => (
                                <ListItem
                                    key={order.id}
                                    model="sale.order"
                                    res_id={order.id}
                                    title={this.getTitle(order)}
                                    description={this.getOrderDescription(order)}
                                    logTitle={_t('Link to Odoo')}
                                    partnerId={partner.id}
                                    hasValue={!!order.requested_date}
                                />
                            ))}
                        </>
                    ) : (
                        <div className="no-records">
                            {_t('No quotations found for this contact.')}
                        </div>
                    )}
                </div>
            </div>
        );
    }
}

SaleOrdersPage.contextType = AppContext;

export default SaleOrdersPage;
