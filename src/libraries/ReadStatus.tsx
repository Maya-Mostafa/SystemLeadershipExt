import * as React from 'react';

export interface IReadStatusProps {
    userPropIds: any;
    listItemID: string;
}

export default function ReadStatus (props: IReadStatusProps){

    // console.log("props.userPropIds", props.userPropIds);
    // console.log("props.listItemID", props.listItemID);

    return(
        <>
            {props.userPropIds.has(props.listItemID) ? 
                <span style={{color: 'green'}}>Yes</span> : 
                <span style={{color: 'red'}}>No</span>
            }
        </>
    )
}