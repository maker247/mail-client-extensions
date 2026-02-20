import Lead from './Lead';
import Task from './Task';
import HelpdeskTicket from './HelpdeskTicket';
import SaleOrder from './SaleOrder';

export default class User {
    id: number;
    name: string;
    email: string;
    company: string;
    leads?: Lead[];
    tasks?: Task[];
    tickets?: HelpdeskTicket[];
    saleOrders?: SaleOrder[];
    recentSaleOrderId?: number;

    constructor(id: number, name: string, email: string, company: string, recentSaleOrderId?: number) {
        this.id = id;
        this.name = name;
        this.email = email;
        this.company = company;
        this.recentSaleOrderId = recentSaleOrderId;
    }

    static getEmptyUser(): User {
        return new User(0, '', '', '');
    }

    // Temporary method to build a mock user for testing
    static getMockUser(): User {
        let u = new User(1, 'Demo User', 'user@demo.inc', 'Demo Company', 0);
        u.saleOrders = [];
        u.tasks = [];
        u.tickets = [];
        u.leads = [];
        return u;
    }
}
