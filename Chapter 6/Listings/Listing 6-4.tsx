/* Section - Dependency imports - packages */
import * as React from 'react';
import { useEffect, useState } from 'react';
import { Icon } from '@fluentui/react';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";

/* Section  - Dependency imports - internal */
import styles from './CloudhadiServicePortal.module.scss';
import { ICloudhadiServicePortalProps } from './ICloudhadiServicePortalProps';
import RequestDetails from './RequestDetails';
import RequestList from './RequestList';



const CloudhadiServicePortal = (props: ICloudhadiServicePortalProps): JSX.Element => {

  /*Section - Operations*/

  // Get the context from props.
  const { context } = props;

  //State variables.
  const [activeTab, setActiveTab] = React.useState('createRequest');
  const [isServiceUser, setIsServiceUser] = useState(false);

  // handle tab click.
  const handleTabChange = (tab: string): void => {
    setActiveTab(tab);
  };

  //Get the groups of logged in users
  const checkServiceUser = async (): Promise<void> => {
    try {
      const sp = spfi().using(SPFx(props.context));
      const groups = await sp.web.currentUser.groups();
      const userGroups = groups.map((g) => g.Title);

      if (userGroups.includes("Service Executives")) { setIsServiceUser(true); }
      else { setIsServiceUser(false); }

    } catch (error) {
      console.log("Error fetching user groups:", error);
    }
  };

  // During component load for the first time.
  useEffect(() => {

    checkServiceUser().catch((error) => {
      console.log("Unhandled error in fetching groups:", error);
    });
  }, []);

  // Defining the tabs
  const tabData = [
    {
      tabId: 'createRequest',
      iconName: 'Glimmer',
      label: 'Create a New Request',
      visible: true
    },
    {
      tabId: 'myRequests',
      iconName: 'ContactList',
      label: 'My Requests',
      visible: true,
    },
    {
      tabId: 'pendingRequests',
      iconName: 'GroupedList',
      label: 'Pending Requests',
      visible: isServiceUser
    }
  ];

  /*Section - return function */
  return (
    <div className={styles.cloudhadiServicePortal}>
      <div className={styles.tabContainer}>
        {tabData.filter((tab) => tab.visible).map((tab) => (
          <div
            key={tab.tabId}
            className={`${styles.tab} ${activeTab === tab.tabId ? styles.activeTab : ''}`}
            onClick={() => handleTabChange(tab.tabId)}
          >
            <Icon iconName={tab.iconName} className={styles.tabIcon} />
            {tab.label}
          </div>
        ))}

      </div>
      <div className={styles.tabContent}>
        {activeTab === 'createRequest' && (
          <RequestDetails context={context} />
        )}
        {activeTab === 'myRequests' && (
          <RequestList context={context} isServiceUser={false} />
        )}
        {isServiceUser && activeTab === 'pendingRequests' && (
          <RequestList context={context} isServiceUser={true} />
        )}
      </div>

    </div>
  );
};
/*Section Export */
export default CloudhadiServicePortal;

