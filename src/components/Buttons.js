import React from 'react';
import styled from 'styled-components';

const Button = styled.button`
  padding: 8px 16px;
  border-radius: 4px;
  font-size: 16px;
  cursor: pointer;
`;

const CreateFileButton = styled(Button)`
  background-color: yellow;
  color: black;
`;

const MessagesButton = styled(Button)`
  background-color: orange;
  color: white;
`;

const Buttons = () => {
  return (
    <div>
      <CreateFileButton>Crear archivo</CreateFileButton>
      <MessagesButton>Mensajes</MessagesButton>
    </div>
  );
};

export default Buttons;
