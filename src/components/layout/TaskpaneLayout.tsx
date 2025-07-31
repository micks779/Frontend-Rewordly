import React from 'react';
import { Stack, makeStyles, IStackStyles } from '@fluentui/react';

// Following Microsoft's Office Add-in Design Guidelines
const useStyles = makeStyles({
  root: {
    width: '100%',
    height: '100%',
    backgroundColor: '#f3f2f1 !important', // Light gray background like in the design
    overflowX: 'hidden',
    overflowY: 'auto',
    position: 'relative',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '16px'
  },
  container: {
    backgroundColor: '#fff',
    borderRadius: '8px',
    boxShadow: '0 4px 6px rgba(0, 0, 0, 0.1)',
    width: '100%',
    height: '100%',
    maxWidth: '100%',
    maxHeight: '100%',
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column'
  },
  header: {
    backgroundColor: '#2b579a', // Microsoft Blue
    padding: '12px 16px',
    width: '100%',
    boxSizing: 'border-box',
    borderTopLeftRadius: '8px',
    borderTopRightRadius: '8px'
  },
  content: {
    width: '100%',
    padding: '16px',
    boxSizing: 'border-box',
    flex: '1 1 auto',
    overflowX: 'hidden',
    overflowY: 'auto'
  }
});

interface TaskpaneLayoutProps {
  children: React.ReactNode;
  title?: string;
  hideHeader?: boolean;
}

const TaskpaneLayout: React.FC<TaskpaneLayoutProps> = ({ 
  children, 
  title = "Rewordly",
  hideHeader = false 
}) => {
  const classes = useStyles();

  return (
    <div className={classes.root}>
      <div className={classes.container}>
        {!hideHeader && (
          <div className={classes.header}>
            <span style={{ color: '#fff', fontSize: '16px', fontWeight: 600 }}>{title}</span>
          </div>
        )}
        <div className={classes.content}>
          {children}
        </div>
      </div>
    </div>
  );
};

export default TaskpaneLayout;
