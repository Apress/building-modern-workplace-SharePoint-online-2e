/*Section - Imports */
import * as React from 'react';
import styles from './CloudhadiServicePortal.module.scss';
import { useEffect, useState } from 'react';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { AssignFrom } from "@pnp/core";
import { IRequestListProps } from "./IRequestListProps";
import { Stack, Checkbox, Link, PrimaryButton, DefaultButton, Dialog, DialogType, DialogFooter } from '@fluentui/react';

const RequestList = (props: IRequestListProps): JSX.Element => {
  /*Section - component body */
  const [listData, setListData] = useState([]);
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [selectedItem, setSelectedItem] = useState(null);
  const [message, setMessage] = useState('');
  const [isSuccess, setIsSuccess] = useState(false);
  const { context, isServiceUser } = props;
  const sp = spfi().using(SPFx(context));
  const spSite = spfi(`${context.pageContext.web.absoluteUrl.split('/sites')[0]}/sites/workplace/`).using(AssignFrom(sp.web));
  const getMyItems = async (): Promise<any[]> => {
   // const workplaceWeb = Web(`https://cloudhadi.sharepoint.com/sites/Workplace`);
    const myItems = await spSite.web.lists.getByTitle(`Service Portal`).items
      .select("Author/EMail", "ID", "Title", "RequestTitle", "Relatedto", "RequestStatus", "RequestDescription")
      .expand("Author")
      .filter(`Author/EMail eq '${context.pageContext.user.email}'`)();
    return myItems;
  }

  const getPendingItems = async (): Promise<any[]> => {
    const newItems = await spSite.web.lists.getByTitle(`Service Portal`).items
      .select("Author/Title", "ID", "Title", "RequestTitle", "Relatedto", "RequestStatus", "RequestDescription")
      .expand("Author")
      .filter(`RequestStatus ne 'Completed' || RequestStatus ne 'Rejected'`)()

    return newItems;
  }


  useEffect(() => {

    const getItems = async (): Promise<void> => {
      try {

        let items: any[];

        if (isServiceUser) {
          items = await getPendingItems();
        }
        else {
          items = await getMyItems();
        }
        const updatedItems = items.map(item => ({
          ...item,
          checked: false
        }));

        setListData(updatedItems);

      } catch (error) {
        console.log("Error fetching pending items:", error);
      }
    };

    getItems().catch((error) => {
      console.log("Unhandled error in fetching items:", error);
    });
  }, [isSuccess]);

  const showDetailsDialog = (item: unknown): void => {
    setSelectedItem(item);
    setIsDialogOpen(true);
  };

  const hideDetailsDialog = (): void => {
    setIsDialogOpen(false);
  };
  const updateStatus = async (updatedStatus: string): Promise<void> => {
    const selectedItems = listData.filter(item => item.checked);
    if (selectedItems.length === 0) {
      setMessage('Please select a request to set status.');
      return;
    }

    try {
      await Promise.all(
        selectedItems.map(async (item) => {
          await sp.web.lists.getByTitle("Service Portal").items.getById(item.ID).update({
            RequestStatus: updatedStatus,
          });
        })
      );

      setMessage('Status updated for selected request/s!');
      setIsSuccess(true);
    } catch (error) {
      setMessage('Status update failed. Please contact IT team');
    }
  };


  const handleCheckboxChange = React.useCallback(
    (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean, itemId?: number) => {
      setListData(prevListData =>
        prevListData.map(item => {
          if (item.ID === itemId) {
            return {
              ...item,
              checked: checked
            };
          }
          return item;
        })
      );
    },
    []
  );
  /*Section - return function */
  return (
    <div className={styles.requestList}>
      <h2 className={styles.title}>{isServiceUser ? `Pending Requests` : 'My Requests'}</h2>
      {isServiceUser && <Stack className={styles.buttonList} horizontal tokens={{ childrenGap: 20 }}>
        <PrimaryButton text="In Progress" onClick={() => updateStatus("In Progress")} className={styles.btnInProgress} />
        <PrimaryButton text="Resolve" onClick={() => updateStatus("Resolved")} className={styles.fluentControl} />
        <PrimaryButton text="Complete" onClick={() => updateStatus("Completed")} className={styles.btnCompleted} />
        <DefaultButton text="Reject" onClick={() => updateStatus("Rejected")} className={styles.fluentControl} />
      </Stack>}
      {message && (
        <div className={isSuccess ? styles.successLabel : styles.errorLabel}>
          <span>{message}</span>
        </div>
      )}
      <table className={styles.table}>
        <thead>
          <tr>
            <th>ID</th>
            <th>Title</th>
            <th>Related To</th>
            {!isServiceUser && <th>Status</th>}
            {isServiceUser && <><th>Requested By</th>
              <th>Action</th></>}
          </tr>
        </thead>
        <tbody>
          {listData.map((item, index) => (
            <tr key={index}>
              <td> <Link onClick={() => showDetailsDialog(item)}>{item.Title}</Link></td>
              <td>{item.RequestTitle}</td>
              <td>{item.Relatedto}</td>
              {!isServiceUser && <td>{item.RequestStatus}</td>}
              {isServiceUser && <><td>{item.Author.Title}</td>
                <td> <Checkbox checked={item.checked} onChange={(ev, checked) => handleCheckboxChange(ev, checked, item.ID)} /></td>
              </>}

            </tr>
          ))}
        </tbody>
      </table>
      {/* Details Dialog */}
      <Dialog
        hidden={!isDialogOpen}
        onDismiss={hideDetailsDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: selectedItem?.Title || '',
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 450 } },
        }}
      >
        {selectedItem && (
          <div>
            <p><strong><em>Request Title:</em></strong> {selectedItem.RequestTitle}</p>
            <p><strong><em>Description:</em></strong> {selectedItem.RequestDescription}</p>
            <p><strong><em>Related To:</em></strong> {selectedItem.Relatedto}</p>
            <p><strong><em>Current Status:</em></strong> {selectedItem.RequestStatus}</p>
          </div>

        )}
        <DialogFooter>
          <PrimaryButton onClick={hideDetailsDialog} text="Close" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default RequestList;

