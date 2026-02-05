import * as React from 'react';
import { IRequestListProps } from "./IRequestListProps";
const RequestList = (props: IRequestListProps) => {
  const {context, isServiceUser } = props;
  return (
    <div>
      <h2>Request List - {context.pageContext.user.displayName}, {isServiceUser}</h2>
    </div>
  );
};

export default RequestList;
