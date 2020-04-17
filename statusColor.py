class StatusColor:
   color="No Status"
   emoji=''
   
def readStatusColor(color):
   status=StatusColor()
   if(color==1):
      return(status) 
   rgb=color[-6:]
   r=int(rgb[:2],16)
   g=int(rgb[2:-2],16)
   b=int(rgb[-2:],16)
   if(r>b and g>b):
      status.color="Yellow"
      status.emoji="🟡"
   elif(g>r):
      status.color="Green"
      status.emoji="🟢"
   elif(r>g and r>b):
      status.color="Red"
      status.emoji="🔴"
   return(status)

def setCode(cell):
   color=readStatusColor(cell.font.color.index)
   return(color)
   
def setEmoji(cell):
   color=readStatusColor(cell.font.color.index)
   return(color.emoji)

def setColor(cell):
   color=readStatusColor(cell.font.color.index)
   return(color.color)
