import * as React from 'react';
import './Logger.css';
import { ContentType, HttpVerb, sendHttpRequest } from '../../../utils/httpRequest';
import AppContext from '../AppContext';
import api from '../../api';
import { Spinner, SpinnerSize, TooltipHost, Dialog, DialogType, DialogFooter, PrimaryButton, DefaultButton } from 'office-ui-fabric-react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faCheck, faEnvelope, faSync } from '@fortawesome/free-solid-svg-icons';
import { OdooTheme } from '../../../utils/Themes';
import { _t } from '../../../utils/Translator';
import PostalMime from 'postal-mime';

//total attachments size threshold in megabytes
const SIZE_THRESHOLD_TOTAL = 40;

//single attachment size threshold in megabytes
const SIZE_THRESHOLD_SINGLE_ELEMENT = 10;

type LoggerProps = {
    resId: number;
    model: string;
    tooltipContent: string;
    partnerId?: number;
    hasValue?: boolean;
};

type LoggerState = {
    logged: number;
    isConfirmDialogOpen: boolean;
};

class Logger extends React.Component<LoggerProps, LoggerState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            logged: 0,
            isConfirmDialogOpen: false,
        };
    }

    private arrayBufferToBase64(buffer) {
        const bytes = new Uint8Array(buffer);
        const chunkSize = 0x8000; // 32KB
        let binary = '';

        for (let i = 0; i < bytes.length; i += chunkSize) {
            binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
        }

        return btoa(binary);
    }

    private logRequest = async (event): Promise<any> => {
        event?.stopPropagation?.();

        if (this.props.hasValue && !this.state.isConfirmDialogOpen) {
            this.setState({ isConfirmDialogOpen: true });
            return;
        }

        this.setState({ logged: 1, isConfirmDialogOpen: false });
        Office.context.mailbox.item.getAsFileAsync(async (result) => {
            if (!result.value && result.error) {
                this.context.showHttpErrorMessage(result.error);
                this.setState({ logged: 0 });
                return;
            }

            const parser = new PostalMime();
            const email = await parser.parse(atob(result.value));
            const doc = new DOMParser().parseFromString(email.html, 'text/html');

            let node: Element = doc.getElementById('appendonsend');
            // Remove the history and only log the most recent message.
            while (node) {
                const next = node.nextElementSibling;
                node.parentNode.removeChild(node);
                node = next;
            }
            const msgHeader = `<div>${_t('From : %(email)s', {
                email: email.from.address,
            })}</div>`;
            doc.body.insertAdjacentHTML('afterbegin', msgHeader);
            const mailId = encodeURIComponent(Office.context.mailbox.item.itemId);
            const mailUrl = `https://outlook.office.com/mail/deeplink/read/${mailId}`;
            const msgFooter = `<br/><div class="text-muted font-italic">${_t(
                'Logged from',
            )} <a href="${mailUrl}" target="_blank">${_t(
                'Outlook Inbox',
            )}</a></div>`;
            doc.body.insertAdjacentHTML('beforeend', msgFooter);

            const totalSize = email.attachments.reduce((sum, attachment) => sum + attachment.content.byteLength, 0);
            if (totalSize > SIZE_THRESHOLD_TOTAL * 1024 * 1024) {
                const warningMessage = _t(
                    'Warning: Attachments could not be logged in Odoo because their total size' +
                    ' exceeded the allowed maximum.',
                    {
                        size: SIZE_THRESHOLD_TOTAL,
                    },
                );
                doc.body.innerHTML += `<div class="text-danger">${warningMessage}</div>`;
                email.attachments = [];
            }

            const standardAttachments = [];
            const oversizedAttachments = [];
            const inlineAttachments = {};
            email.attachments.forEach((attachment) => {
                if (attachment.disposition === 'inline') {
                    inlineAttachments[attachment.contentId] = attachment;
                } else if (attachment.content.byteLength > SIZE_THRESHOLD_SINGLE_ELEMENT * 1024 * 1024) {
                    oversizedAttachments.push(attachment.filename);
                } else {
                    standardAttachments.push([attachment.filename, this.arrayBufferToBase64(attachment.content)]);
                }
            });

            if (oversizedAttachments.length > 0) {
                const warningMessage = _t(
                    'Warning: Could not fetch the attachments %(attachments)s as their sizes are bigger then the maximum size of %(size)sMB per each attachment.',
                    {
                        attachments: oversizedAttachments.join(', '),
                        size: SIZE_THRESHOLD_SINGLE_ELEMENT,
                    },
                );
                doc.body.innerHTML += `<div class="text-danger">${warningMessage}</div>`;
            }

            const imageElements = Array.from(doc.getElementsByTagName('img')).filter((img) =>
                img.getAttribute('src')?.startsWith('cid:'),
            );
            imageElements.forEach((element) => {
                const attachment = inlineAttachments[`<${element.src.replace(/^cid:/, '')}>`];
                if (attachment?.content.byteLength > SIZE_THRESHOLD_SINGLE_ELEMENT * 1024 * 1024) {
                    element.setAttribute(
                        'alt',
                        _t('Could not display image %(attachmentName)s, size is over limit', {
                            attachmentName: attachment.filename,
                        }),
                    );
                } else if (attachment) {
                    const fileExtension = attachment.filename.split('.')[1];
                    element.setAttribute(
                        'src',
                        `data:image/${fileExtension};base64, ${this.arrayBufferToBase64(attachment.content)}`,
                    );
                }
            });

            const requestJson = {
                res_id: this.props.resId,
                model: this.props.model,
                message: doc.documentElement.outerHTML,
                attachments: standardAttachments,
                ...(this.props.model === 'sale.order' ? {
                    partner_id: this.props.partnerId,
                    requested_mail_id: mailId,
                    requested_at: ((): string => {
                        const d = Office.context.mailbox.item.dateTimeCreated;
                        const year = d.getUTCFullYear();
                        const month = ('0' + (d.getUTCMonth() + 1)).slice(-2);
                        const day = ('0' + d.getUTCDate()).slice(-2);
                        const hours = ('0' + d.getUTCHours()).slice(-2);
                        const minutes = ('0' + d.getUTCMinutes()).slice(-2);
                        const seconds = ('0' + d.getUTCSeconds()).slice(-2);
                        return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
                    })()
                } : {})
            };

            const logRequest = sendHttpRequest(
                HttpVerb.POST,
                api.baseURL + (this.props.model === 'sale.order' ? api.logSaleOrder : api.logSingleMail),
                ContentType.Json,
                this.context.getConnectionToken(),
                requestJson,
                true,
            );
            logRequest.promise
                .then((response) => {
                    const { result } = JSON.parse(response);
                    if (result.error) {
                        this.setState({ logged: 0 });
                        this.context.showValidationErrorMessage(result.error);
                        return;
                    } else {
                        this.setState({ logged: 2 });
                    }
                })
                .catch((error) => {
                    this.context.showHttpErrorMessage(error);
                    this.setState({ logged: 0 });
                });
        });
    };

    render() {
        let logContainer = null;
        switch (this.state.logged) {
            case 0:
                logContainer = (
                    <div className="log-container">
                        <TooltipHost content={this.props.tooltipContent}>
                            <div className={`odoo-secondary-button log-button ${this.props.hasValue ? 'has-value' : ''}`} onClick={this.logRequest}>
                                {this.props.model === 'sale.order' && !this.props.hasValue ? (
                                    <svg viewBox="0 0 2134 2134" className="custom-log-icon">
                                        <path d="M801.864,814.773l-549.544,-0c-90.994,-0 -164.82,73.826 -164.82,164.82l0,65.241c0,25.41 11.675,49.103 31.591,64.898c105.073,82.067 139.754,111.254 595.412,441.924c57.687,41.892 172.374,143.531 252.037,142.157c79.663,1.374 194.35,-100.266 252.037,-142.157c145.286,-105.433 247.773,-180.182 323.39,-235.766c-24.082,2.537 -48.531,3.838 -73.281,3.838c-317.233,0 -584.937,-213.748 -666.823,-504.956Z" fill="currentColor" />
                                        <path d="M1812.273,1250.173c13.392,-10.645 33.307,-0.687 33.307,16.139l0,702.202c0,90.994 -73.826,164.82 -164.82,164.82l-1428.44,0c-90.994,0 -164.82,-73.826 -164.82,-164.82l0,-701.859c0,-17.169 19.572,-26.783 33.307,-16.139c76.916,59.747 178.898,135.633 529.141,390.074c72.452,52.88 194.694,164.133 316.592,163.447c122.585,1.03 247.23,-112.627 316.935,-163.447c350.243,-254.441 451.882,-330.67 528.798,-390.417Z" fill="currentColor" />
                                        <path d="M1468.687,19.598c-335.621,0 -607.548,271.927 -607.548,607.548c0,335.621 271.927,607.548 607.548,607.548c335.621,0 607.548,-271.927 607.548,-607.548c0,-335.621 -271.927,-607.548 -607.548,-607.548Zm226.581,766.785l-48.996,61.245c-7.433,9.294 -18.709,14.714 -30.61,14.714c-8.899,0 -17.537,-3.03 -24.486,-8.589l-164.136,-121.804c-23.216,-18.587 -36.751,-46.767 -36.747,-76.507l0,-381.065c0,-21.503 17.694,-39.197 39.197,-39.197l78.393,0c21.503,0 39.197,17.694 39.197,39.197l0,352.77l142.088,104.116c9.292,7.434 14.711,18.708 14.711,30.607c0,8.91 -3.038,17.56 -8.611,24.513Z" fill="currentColor" />
                                    </svg>
                                ) : (
                                    <FontAwesomeIcon
                                        icon={this.props.model === 'sale.order' ? faSync : faEnvelope}
                                    />
                                )}
                            </div>
                        </TooltipHost>
                    </div>
                );
                break;
            case 1:
                logContainer = (
                    <div className="log-container">
                        <div>
                            <Spinner theme={OdooTheme} size={SpinnerSize.medium} />
                        </div>
                    </div>
                );
                break;
            case 2:
                logContainer = (
                    <div className="log-container">
                        <div className="logged-text">
                            <FontAwesomeIcon icon={faCheck} color={'green'} />
                        </div>
                    </div>
                );
                break;
        }

        const dialogContentProps = {
            type: DialogType.normal,
            title: _t('Update Confirmation'),
            closeButtonAriaLabel: _t('Close'),
            subText: _t('Are you sure you want to update the requested time?'),
        };

        return (
            <>
                {logContainer}
                <Dialog
                    hidden={!this.state.isConfirmDialogOpen}
                    onDismiss={() => this.setState({ isConfirmDialogOpen: false })}
                    dialogContentProps={dialogContentProps}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={() => this.logRequest(null)} text={_t('Yes')} />
                        <DefaultButton onClick={() => this.setState({ isConfirmDialogOpen: false })} text={_t('No')} />
                    </DialogFooter>
                </Dialog>
            </>
        );
    }
}

Logger.contextType = AppContext;

export default Logger;
