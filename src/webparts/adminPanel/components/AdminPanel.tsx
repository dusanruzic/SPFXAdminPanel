import * as React from 'react';
import styles from './AdminPanel.module.scss';
import { IAdminPanelProps } from './IAdminPanelProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import SharePointService from '../../../services/SharePoint/SharePointService';
import {UserGroup} from './UserGroup';

export default class AdminPanel extends React.Component<IAdminPanelProps, {}> {
  public render(): React.ReactElement<IAdminPanelProps> {
    return (
      <div className={ styles.adminPanel }>
        <h1 style={{textAlign: 'center' }}>
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
          <UserGroup name="User"></UserGroup>

        </PivotItem>

        <PivotItem headerText="SoftwareDeveloper group">
        <UserGroup name="SoftwareDeveloper"></UserGroup>
        </PivotItem>

        <PivotItem headerText="Admin group">
        <UserGroup name="Admin"></UserGroup>
        </PivotItem>   
    
        
      </Pivot>  
      </div>
    );
  }
}
