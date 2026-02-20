import * as React from 'react';
import User from '../../../../classes/User';
import '../../Contact/ContactList/ContactListItem/ContactListItem.css';
import { faEnvelope, faBuilding } from '@fortawesome/free-solid-svg-icons';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { TooltipHost, TooltipOverflowMode } from 'office-ui-fabric-react';

type UserCardProps = {
    user: User;
};

class UserCard extends React.Component<UserCardProps> {

    getInitials(name: string) {
        if (!name) return 'U';
        const parts = name.split(' ');
        if (parts.length > 1) {
            return parts[0][0].toUpperCase() + parts[1][0].toUpperCase();
        }
        return name[0].toUpperCase();
    }

    render() {
        const { user } = this.props;

        const emailDiv = user.email && (
            <div className="contact-email">
                <FontAwesomeIcon icon={faEnvelope} className="contact-info-icon fa-fw" />
                <div className="text-truncate">
                    <TooltipHost content={user.email} overflowMode={TooltipOverflowMode.Parent}>
                        {user.email}
                    </TooltipHost>
                </div>
            </div>
        );

        const companyDiv = user.company && (
            <div className="contact-email">
                <FontAwesomeIcon icon={faBuilding} className="contact-info-icon fa-fw" />
                <div className="text-truncate">
                    <TooltipHost content={user.company} overflowMode={TooltipOverflowMode.Parent}>
                        {user.company}
                    </TooltipHost>
                </div>
            </div>
        );

        return (
            <div className="contact-container" style={{ borderBottom: '1px solid #e1e1e1', cursor: 'default' }}>
                <div className="contact-card">
                    <div data-initials={this.getInitials(user.name)} />
                    <div className="contact-info">
                        <div className="contact-name">{user.name || 'Unknown User'}</div>
                        {emailDiv}
                        {companyDiv}
                    </div>
                </div>
            </div>
        );
    }
}

export default UserCard;
