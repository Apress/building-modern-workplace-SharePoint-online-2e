import * as React from 'react';
import styles from './CloudhadiServicePortal.module.scss';
import {ICloudhadiServicePortalProps} from './ICloudhadiServicePortalProps';

const CloudhadiServicePortal = (props: ICloudhadiServicePortalProps) => {
  return (
    <div className={styles.welcome}>
      Service Portal - Welcome {props.userDisplayName}
    </div>
  );
};

export default CloudhadiServicePortal;
