import React, { useEffect, useState } from 'react';
import './App.css';
import { Login, Person } from '@microsoft/mgt-react';
import { PersonViewType} from '@microsoft/mgt';
import { useGet, useIsSignedIn } from './mgt';
import { buildColumns, DetailsList, IColumn, SelectionMode, Spinner, SpinnerSize } from '@fluentui/react';
import { Message, User } from '@microsoft/microsoft-graph-types'

function App() {

  const [isSignedIn] = useIsSignedIn();
  const [user, userLoading ] = useGet<User>('/me');

  return (
    <div className="App">
      <header>
        <Login></Login>
      </header>
      {isSignedIn &&
        <div>
          { !userLoading && <div>
              <h3> Hello {user?.displayName},</h3>
              <h5> Here are your messages! </h5>
            </div>
          }
          <Mail></Mail>
        </div>
      }
    </div>
  );
}

function Mail() {

  let [messages] = useGet('/me/messages');

  let [sortedMessages, setSortedMessages] = useState<any[]>(messages);
  let [columns, setColumns] = useState<IColumn[]>();

  useEffect(() => {
    if (messages && messages.value && messages.value.length) {
      const items = messages?.value?.map((m: Message) => {
        return {
          key: m.id,
          from: m.sender?.emailAddress?.address,
          subject: m.subject,
          preview: m.bodyPreview
        }
      })

      setSortedMessages(items as any[]);

      setColumns(buildColumns(items as any[]))
    }

  }, [messages]);

  if (sortedMessages && sortedMessages.length) {

    const renderItemColumn = (item?: any, index?: number, column?: IColumn) => {

      if (!item || !column) {
        return <div></div>
      }

      const fieldContent = item[column.fieldName as string] as string;
    
      switch (column.key) {
        case 'from':
          return <Person personQuery={fieldContent} 
            view={PersonViewType.oneline} >
          </Person>
    
        default:
          return <span>{fieldContent}</span>;
      }
    }

    const onColumnClick = (event?: React.MouseEvent<HTMLElement>, column?: IColumn): void => {
      if (!sortedMessages || !column) {
        return;
      }
      
      let isSortedDescending = column.isSortedDescending;
  
      // If we've sorted this column, flip it.
      if (column.isSorted) {
        isSortedDescending = !isSortedDescending;
      }
  
      // Sort the items.
      setSortedMessages(_copyAndSort(sortedMessages, column.fieldName!, isSortedDescending));
  
      setColumns(columns?.map(col => {
          col.isSorted = col.key === column.key;
  
          if (col.isSorted) {
            col.isSortedDescending = isSortedDescending;
          }
  
          return col;
        }));
    };

    return <div className="MessagesMain">
        <DetailsList 
          onColumnHeaderClick={onColumnClick}
          onRenderItemColumn={renderItemColumn}
          selectionMode={SelectionMode.none} 
          items={sortedMessages as any[]} 
          columns={columns} ></DetailsList>
      </div>
  }

  return <div>No messages</div>;
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}


export default App;
