import * as React from 'react';
import styles from './AdminPanel.module.scss';
import { IAdminPanelProps } from './IAdminPanelProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import SharePointService from '../../../services/SharePoint/SharePointService';
import {Users} from './Users';
import {SoftwareDevelopers} from './SoftwareDevelopers';
import {Admins} from './Admins';

export default class AdminPanel extends React.Component<IAdminPanelProps, {}> {
  public render(): React.ReactElement<IAdminPanelProps> {
    return (
      <div className={ styles.adminPanel }>
        <h1>
          Admin panel
        </h1>
        <Pivot aria-label="All users">
        <PivotItem
          headerText="User group"
          headerButtonProps={{
            'data-order': 1,
            'data-title': 'Users',
          }}
        >
          <Users></Users>

        </PivotItem>

        <PivotItem headerText="Software developer group">
        <SoftwareDevelopers></SoftwareDevelopers>
        </PivotItem>

        <PivotItem headerText="Admin group">
        <Admins></Admins>
        </PivotItem>        
        
      </Pivot>  
      </div>
    );
  }
}
