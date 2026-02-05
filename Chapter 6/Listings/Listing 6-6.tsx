/*Section - imports */

import * as React from 'react';
import { useState } from 'react';
import { Stack, TextField, Dropdown, PrimaryButton, DefaultButton, IDropdownOption } from '@fluentui/react';
import styles from './CloudhadiServicePortal.module.scss';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { AssignFrom } from "@pnp/core";
import { IRequestDetailsProps } from "./IRequestDetailsProps";

const RequestDetails = (props: IRequestDetailsProps): JSX.Element => {

  /*Section - component body */

  const [requestTitle, setRequestTitle] = useState('');
  const [requestDescription, setRequestDescription] = useState('');
  const [relatedTo, setRelatedTo] = useState('');

  const [message, setMessage] = useState('');
  const [IsSuccess, setIsSuccess] = useState(false);
  const { context } = props;
  const sp = spfi().using(SPFx(context));
  const spSite = spfi(`${context.pageContext.web.absoluteUrl.split('/sites')[0]}/sites/workplace/`).using(AssignFrom(sp.web));
  const getRandomInt = (min: number, max: number): number => {
    min = Math.ceil(min);
    max = Math.floor(max);
    return Math.floor(Math.random() * (max - min + 1)) + min;
  };

  const handleReset = (): void => {
    setRequestTitle('');
    setRequestDescription('');
    setRelatedTo('');
  };


  const handleSubmit = async (): Promise<void> => {
    if (!requestTitle || !requestDescription || !relatedTo) {
      setMessage('Please fill in all the required fields.');
      return;
    }
    try {

      const reqID = `CSR${getRandomInt(10000, 99999)}`;
      await  spSite.web.lists.getByTitle(`Service Portal`)
        .items.add({
          Title: reqID,
          RequestTitle: requestTitle,
          RequestDescription: requestDescription,
          Relatedto: relatedTo,
          RequestStatus: "New"
        });
      handleReset();
      setIsSuccess(true);
      setMessage(`Service request Created! ${reqID}`);
    }
    catch (Ex) {
      setMessage('Service request creation failed. Please contact IT team');
    }

  };

  const handleRelatedToChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      setRelatedTo(option.key.toString());
    }
  };

  /*Section - return function */
  return (
    <div className={styles.container}>
      <h2 className={styles.title}>New Service Request</h2>
      <div className={styles.formGrid}>
        <Stack tokens={{ childrenGap: 15 }}>
          <TextField label="Request Title"
            value={requestTitle}
            onChange={(event, newValue) => setRequestTitle(newValue || '')}
            required className={styles.fluentControl} />
          <TextField label="Request Description"
            value={requestDescription}
            onChange={(event, newValue) => setRequestDescription(newValue || '')}
            multiline rows={4} required className={styles.fluentControl} />
          <Dropdown
            label="Related to"
            defaultSelectedKey={relatedTo}
            onChange={handleRelatedToChange}
            options={[
              { key: 'Access', text: 'Access' },
              { key: 'Materials', text: 'Materials' },
              { key: 'Equipment', text: 'Equipment' },
              { key: 'General', text: 'General' }
            ]}
            required
            className={styles.fluentControl}
          />
          <Stack className={styles.buttonContainer} horizontal tokens={{ childrenGap: 20 }}>
            <PrimaryButton text="Submit" onClick={handleSubmit} className={styles.fluentControl} />
            <DefaultButton text="Reset" onClick={handleReset} className={styles.fluentControl} />
          </Stack>
        </Stack>
        {message && (
          <div className={IsSuccess ? styles.successLabel : styles.errorLabel}>
            <span>{message}</span>
          </div>
        )}
      </div>
    </div>
  );
};

export default RequestDetails;
