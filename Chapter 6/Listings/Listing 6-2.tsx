import * as React from 'react';
import { IRequestDetailsProps } from "./IRequestDetailsProps";
const RequestDetails = (props: IRequestDetailsProps) => {

  const {context} = props;
  return (
    <div>
      <h2>Request Details {context.pageContext.user.displayName}</h2>
    </div>
  );
};

export default RequestDetails;
