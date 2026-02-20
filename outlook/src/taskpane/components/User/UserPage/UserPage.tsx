import * as React from 'react';
import User from '../../../../classes/User';
import AppContext from '../../AppContext';
import UserCard from '../UserCard/UserCard';

// Temporary mock of the Partner Sections since they expect a Partner object currently.
// Eventually these should be refactored or wrapped to accept a User/generic interface.
import SectionSaleOrders from '../../SectionSaleOrders/SectionSaleOrders';
import SectionLeads from '../../SectionLeads/SectionLeads';
import SectionTasks from '../../SectionTasks/SectionTasks';
import SectionTickets from '../../SectionTickets/SectionTickets';
import Partner from '../../../../classes/Partner';

type UserPageProps = {
    user: User;
    onSeeAllSaleOrders: (partner?: Partner) => void;
};

type UserPageState = {
    user: User;
    isLoading: boolean;
};

class UserPage extends React.Component<UserPageProps, UserPageState> {
    constructor(props, context) {
        super(props, context);
        console.log(props, 'UserPage');
        this.state = {
            user: props.user,
            isLoading: false
        };
    }

    getPartnerAdapter(): Partner {
        const p = new Partner();
        p.id = this.props.user.id;
        p.name = this.props.user.name;
        p.email = this.props.user.email;
        p.saleOrders = this.props.user.saleOrders;
        p.leads = this.props.user.leads;
        p.tasks = this.props.user.tasks;
        p.tickets = this.props.user.tickets;
        p.currentSaleOrderId = this.props.user.recentSaleOrderId;
        return p;
    }

    render() {
        if (this.state.isLoading) {
            return <div>Loading User Data...</div>; // simple loader 
        }

        const partnerAdapter = this.getPartnerAdapter();

        const leadsList = partnerAdapter.leads && (
            <SectionLeads partner={partnerAdapter} canCreatePartner={false} />
        );

        const tasksList = partnerAdapter.tasks && (
            <SectionTasks
                partner={partnerAdapter}
                canCreatePartner={false}
                canCreateProject={false}
            />
        );

        const ticketsList = partnerAdapter.tickets && (
            <SectionTickets partner={partnerAdapter} canCreatePartner={false} />
        );

        const saleOrdersList = partnerAdapter.saleOrders && (
            <SectionSaleOrders
                partner={partnerAdapter}
                canCreatePartner={false}
                onSeeAll={() => this.props.onSeeAllSaleOrders(partnerAdapter)}
                isUserPage={true}
                recentSaleOrderId={this.props.user.recentSaleOrderId}
            />
        );

        return (
            <div className="user-page">
                <div className="section-card">
                    <UserCard user={this.props.user} />
                </div>

                {leadsList}
                {saleOrdersList}
                {tasksList}
                {ticketsList}
            </div>
        );
    }
}

UserPage.contextType = AppContext;

export default UserPage;
