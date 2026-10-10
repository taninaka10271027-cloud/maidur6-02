from PIL import Image, ImageDraw, ImageFont
Image.MAX_IMAGE_PIXELS=None
import glob
t=Image.open(glob.glob('img/T-*.png')[0]); W,H=t.size
def cr(b): return t.crop((int(b[0]*W),int(b[1]*H),int(b[2]*W),int(b[3]*H)))
a=cr((0.14,0.37,0.31,0.505)); b=cr((0.44,0.36,0.57,0.48)); c=cr((0.80,0.30,0.985,0.47))
font=ImageFont.truetype('/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf',44)
def lab(im,s):
  d=ImageDraw.Draw(im); d.rectangle((0,0,im.width-1,im.height-1),outline='black',width=4)
  w=d.textlength(s,font=font); d.rectangle((0,0,w+24,62),fill='white',outline='black',width=3); d.text((12,8),s,fill='black',font=font)
  return im
h=900
b=b.resize((int(b.width*h/b.height),h)); c=c.resize((int(c.width*h/c.height),h))
wbot=b.width+c.width+30
a=a.resize((wbot,int(a.height*wbot/a.width)))
a=lab(a,'(A) 受電所　400kVA×3台・CET150sq×9条'); b=lab(b,'(B) 整備格納庫　400kVA×3台・9条'); c=lab(c,'(C) 補給倉庫棟　400kVA×2台・6条')
can=Image.new('RGB',(wbot,a.height+h+30),'white'); can.paste(a,(0,0)); can.paste(b,(0,a.height+30)); can.paste(c,(b.width+30,a.height+30))
can.thumbnail((2000,2400)); can.save('out/fig1_temp.png'); print(can.size)
