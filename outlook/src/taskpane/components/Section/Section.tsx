import * as React from 'react';
import Partner from '../../../classes/Partner';
import { ContentType, HttpVerb, sendHttpRequest } from '../../../utils/httpRequest';
import { _t } from '../../../utils/Translator';
import './Section.css';
import CollapseSection from '../CollapseSection/CollapseSection';
import ListItem from '../ListItem/ListItem';
import api from '../../api';
import AppContext from '../AppContext';

type SectionAbstractProps = {
    className?: string;
    records: any[];
    partner: Partner;
    canCreatePartner: boolean;

    // Odoo Record creation
    model: string;
    // endpoint used to create the record
    odooEndpointCreateRecord: string;
    // name of the key returned by Odoo containing the record ID
    odooRecordIdName: string;
    // Odoo action name used for the redirection
    odooRedirectAction: string;
    // Event when we click on the "+" button to create a new record
    // Can be intercepted to give additional values before creating the record
    // (e.g.: Search a project and add the project ID before creating a task)
    onClickCreate?: (callback: (any?) => void) => void;

    // Messages
    title: string;
    titleCount: string;
    msgNoPartner: string;
    msgNoPartnerNoAccess: string;
    msgNoRecord: string;
    msgLogEmail: string;
    getRecordDescription: (any) => string;
    getRecordHasValue?: (any) => boolean;
    recordCount?: number;
    recentSaleOrderId?: number;
};

type SectionAbstractState = {
    records: any[];
    isCollapsed: boolean;
};

/**
 * Section Component used to display the list of leads, tasks, tickets... Allow to create
 * the record, to log the email on the record or to hide them.
 */
class Section extends React.Component<SectionAbstractProps, SectionAbstractState> {
    constructor(props, context) {
        super(props, context);
        console.log(props, 'Section');
        const isCollapsed = !props.records || !props.records.length;
        this.state = { records: this.props.records, isCollapsed: isCollapsed };
    }

    private onClickCreate = () => {
        if (this.props.onClickCreate) {
            this.props.onClickCreate(this.createRecordRequest);
        } else {
            this.createRecordRequest();
        }
    };

    private createRecordRequest = (additionnalValues?) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, async (result) => {
            // Remove the history and only log the most recent message.
            const message = result.value.split('<div id="x_appendonsend"></div>')[0];
            const subject = Office.context.mailbox.item.subject;

            const requestJson = Object.assign(
                {
                    partner_id: this.props.partner.id,
                    email_body: message,
                    email_subject: subject,
                },
                additionnalValues || {},
            );

            let response = null;
            try {
                response = await sendHttpRequest(
                    HttpVerb.POST,
                    api.baseURL + this.props.odooEndpointCreateRecord,
                    ContentType.Json,
                    this.context.getConnectionToken(),
                    requestJson,
                    true,
                ).promise;
            } catch (error) {
                this.context.showHttpErrorMessage(error);
                return;
            }

            const parsed = JSON.parse(response);
            if (parsed['error'] || (parsed['result'] && parsed['result']['error'])) {
                const errorMessage = parsed['error']?.message || parsed['result']?.error;
                if (errorMessage) {
                    this.context.showValidationErrorMessage(errorMessage);
                } else {
                    this.context.showTopBarMessage();
                }
                return;
            }
            const cids = this.context.getUserCompaniesString();
            const recordId = parsed.result[this.props.odooRecordIdName];
            const url = `${api.baseURL}/web#action=${this.props.odooRedirectAction}&id=${recordId}&model=${this.props.model}&view_type=form${cids}`;
            window.open(url);
        });
    };

    private getSection = () => {
        if (!this.props.partner.isAddedToDatabase()) {
            return (
                <div className="list-text">
                    {_t(this.props.canCreatePartner ? this.props.msgNoPartner : this.props.msgNoPartnerNoAccess)}
                </div>
            );
        } else if (this.state.records.length > 0) {
            return (
                <div className="section-content">
                    {this.props.recentSaleOrderId && <div className="recent">Recent</div>}
                    {this.state.records.map((record) => (
                        <ListItem
                            model={this.props.model}
                            res_id={record.id}
                            key={record.id}
                            title={record.name}
                            description={this.props.getRecordDescription(record)}
                            logTitle={_t(this.props.msgLogEmail)}
                            partnerId={this.props.partner.id}
                            hasValue={this.props.getRecordHasValue?.(record)}
                        />
                    ))}
                </div>
            );
        }
        return <div className="list-text">{_t(this.props.msgNoRecord)}</div>;
    };

    render() {
        const recordCount = this.props.recordCount !== undefined ? this.props.recordCount : (this.state.records && this.state.records.length);
        const title = this.state.records || this.props.recordCount !== undefined
            ? _t(this.props.titleCount, { count: recordCount.toString() })
            : _t(this.props.title);

        return (
            <CollapseSection
                className={this.props.className}
                isCollapsed={this.state.isCollapsed}
                title={title}
                hasAddButton={this.props.partner.isAddedToDatabase() && !!this.props.odooEndpointCreateRecord}
                onAddButtonClick={this.onClickCreate}>
                {this.getSection()}
                {this.props.children}
            </CollapseSection>
        );
    }
}

Section.contextType = AppContext;
export default Section;
