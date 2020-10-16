export let UDOC:any = {};
	
	UDOC.G = {
		concat : function(p:any,r:any) {
			for(var i=0; i<r.cmds.length; i++) p.cmds.push(r.cmds[i]);
			for(var i=0; i<r.crds.length; i++) p.crds.push(r.crds[i]);
		},
		getBB  : function(ps:any) {
			var x0=1e99, y0=1e99, x1=-x0, y1=-y0;
			for(var i=0; i<ps.length; i+=2) {  var x=ps[i],y=ps[i+1];  if(x<x0)x0=x; else if(x>x1)x1=x;  if(y<y0)y0=y;  else if(y>y1)y1=y;  }
			return [x0,y0,x1,y1];
		},
		rectToPath: function(r:any) {  return  {cmds:["M","L","L","L","Z"],crds:[r[0],r[1],r[2],r[1], r[2],r[3],r[0],r[3]]};  },
		// a inside b
		insideBox: function(a:any,b:any) {  return b[0]<=a[0] && b[1]<=a[1] && a[2]<=b[2] && a[3]<=b[3];   },
		isBox : function(p:any, bb:any) {
			var sameCrd8 = function(pcrd:any, crds:any) {
				for(var o=0; o<8; o+=2) {  var eq = true;  for(var j=0; j<8; j++) if(Math.abs(crds[j]-pcrd[(j+o)&7])>=2) {  eq = false;  break;  }    if(eq) return true;  }
				return false;
			};
			if(p.cmds.length>10) return false;
			var cmds=p.cmds.join(""), crds=p.crds;
			var sameRect = false;
			if((cmds=="MLLLZ"  && crds.length== 8) 
			 ||(cmds=="MLLLLZ" && crds.length==10) ) {
				if(crds.length==10) crds=crds.slice(0,8);
				var x0=bb[0],y0=bb[1],x1=bb[2],y1=bb[3];
				if(!sameRect) sameRect = sameCrd8(crds, [x0,y0,x1,y0,x1,y1,x0,y1]);
				if(!sameRect) sameRect = sameCrd8(crds, [x0,y1,x1,y1,x1,y0,x0,y0]);
			}
			return sameRect;
		},
		boxArea: function(a:any) {  var w=a[2]-a[0], h=a[3]-a[1];  return w*h;  },
		newPath: function(gst:any    ) {  gst.pth = {cmds:[], crds:[]};  },
		moveTo : function(gst:any,x:any,y:any) {  var p=UDOC.M.multPoint(gst.ctm,[x,y]);  //if(gst.cpos[0]==p[0] && gst.cpos[1]==p[1]) return;
										gst.pth.cmds.push("M");  gst.pth.crds.push(p[0],p[1]);  gst.cpos = p;  },
		lineTo : function(gst:any,x:any,y:any) {  var p=UDOC.M.multPoint(gst.ctm,[x,y]);  if(gst.cpos[0]==p[0] && gst.cpos[1]==p[1]) return;
										gst.pth.cmds.push("L");  gst.pth.crds.push(p[0],p[1]);  gst.cpos = p;  },
		curveTo: function(gst:any,x1:any,y1:any,x2:any,y2:any,x3:any,y3:any) {   var p;  
			p=UDOC.M.multPoint(gst.ctm,[x1,y1]);  x1=p[0];  y1=p[1];
			p=UDOC.M.multPoint(gst.ctm,[x2,y2]);  x2=p[0];  y2=p[1];
			p=UDOC.M.multPoint(gst.ctm,[x3,y3]);  x3=p[0];  y3=p[1];  gst.cpos = p;
			gst.pth.cmds.push("C");  
			gst.pth.crds.push(x1,y1,x2,y2,x3,y3);  
		},
		closePath: function(gst:any  ) {  gst.pth.cmds.push("Z");  },
		arc : function(gst:any,x:any,y:any,r:any,a0:any,a1:any, neg:any) {
			
			// circle from a0 counter-clock-wise to a1
			if(neg) while(a1>a0) a1-=2*Math.PI;
			else    while(a1<a0) a1+=2*Math.PI;
			var th = (a1-a0)/4;
			
			var x0 = Math.cos(th/2), y0 = -Math.sin(th/2);
			var x1 = (4-x0)/3, y1 = y0==0 ? y0 : (1-x0)*(3-x0)/(3*y0);
			var x2 = x1, y2 = -y1;
			var x3 = x0, y3 = -y0;
			
			var p0 = [x0,y0], p1 = [x1,y1], p2 = [x2,y2], p3 = [x3,y3];
			
			var pth = {cmds:[(gst.pth.cmds.length==0)?"M":"L","C","C","C","C"], crds:[x0,y0,x1,y1,x2,y2,x3,y3]};
			
			var rot = [1,0,0,1,0,0];  UDOC.M.rotate(rot,-th);
			
			for(var i=0; i<3; i++) {
				p1 = UDOC.M.multPoint(rot,p1);  p2 = UDOC.M.multPoint(rot,p2);  p3 = UDOC.M.multPoint(rot,p3);
				pth.crds.push(p1[0],p1[1],p2[0],p2[1],p3[0],p3[1]);
			}
			
			var sc = [r,0,0,r,x,y];  
			UDOC.M.rotate(rot, -a0+th/2);  UDOC.M.concat(rot, sc);  UDOC.M.multArray(rot, pth.crds);
			UDOC.M.multArray(gst.ctm, pth.crds);
			
			UDOC.G.concat(gst.pth, pth);
			var y:any=pth.crds.pop();  x=pth.crds.pop();
			gst.cpos = [x,y];
		},
		toPoly : function(p:any) {
			if(p.cmds[0]!="M" || p.cmds[p.cmds.length-1]!="Z") return null;
			for(var i=1; i<p.cmds.length-1; i++) if(p.cmds[i]!="L") return null;
			var out = [], cl = p.crds.length;
			if(p.crds[0]==p.crds[cl-2] && p.crds[1]==p.crds[cl-1]) cl-=2;
			for(var i=0; i<cl; i+=2) out.push([p.crds[i],p.crds[i+1]]);
			if(UDOC.G.polyArea(p.crds)<0) out.reverse();
			return out;
		},
		fromPoly : function(p:any) {
			var o:any = {cmds:[],crds:[]};
			for(var i=0; i<p.length; i++) { o.crds.push(p[i][0], p[i][1]);  o.cmds.push(i==0?"M":"L");  }
			o.cmds.push("Z");
			return o;
		},
		polyArea : function(p:any) {
			if(p.length <6) return 0;
			var l = p.length - 2;
			var sum = (p[0]-p[l]) * (p[l+1]+p[1]);
			for(var i=0; i<l; i+=2)
				sum += (p[i+2]-p[i]) * (p[i+1]+p[i+3]);
			return - sum * 0.5;
		},
		polyClip : function(p0:any, p1:any) {  // p0 clipped by p1
            var cp1:any, cp2:any, s:any, e:any;
            var inside = function (p:any) {
                return (cp2[0]-cp1[0])*(p[1]-cp1[1]) > (cp2[1]-cp1[1])*(p[0]-cp1[0]);
            };
            var isc = function () {
                var dc = [ cp1[0] - cp2[0], cp1[1] - cp2[1] ],
                    dp = [ s[0] - e[0], s[1] - e[1] ],
                    n1 = cp1[0] * cp2[1] - cp1[1] * cp2[0],
                    n2 = s[0] * e[1] - s[1] * e[0], 
                    n3 = 1.0 / (dc[0] * dp[1] - dc[1] * dp[0]);
                return [(n1*dp[0] - n2*dc[0]) * n3, (n1*dp[1] - n2*dc[1]) * n3];
            };
            var out = p0;
            cp1 = p1[p1.length-1];
            for (let j in p1) {
                var cp2 = p1[j];
                var inp = out;
                out = [];
                s = inp[inp.length - 1]; //last on the input list
                for (let i in inp) {
                    var e = inp[i];
                    if (inside(e)) {
                        if (!inside(s)) {
                            out.push(isc());
                        }
                        out.push(e);
                    }
                    else if (inside(s)) {
                        out.push(isc());
                    }
                    s = e;
                }
                cp1 = cp2;
            }
            return out
        }
	}
	UDOC.M = {
		getScale : function(m:any) {  return Math.sqrt(Math.abs(m[0]*m[3]-m[1]*m[2]));  },
		translate: function(m:any,x:any,y:any) {  UDOC.M.concat(m, [1,0,0,1,x,y]);  },
		rotate   : function(m:any,a:any  ) {  UDOC.M.concat(m, [Math.cos(a), -Math.sin(a), Math.sin(a), Math.cos(a),0,0]);  },
		scale    : function(m:any,x:any,y:any) {  UDOC.M.concat(m, [x,0,0,y,0,0]);  },
		concat   : function(m:any,w:any  ) {  
			var a=m[0],b=m[1],c=m[2],d=m[3],tx=m[4],ty=m[5];
			m[0] = (a *w[0])+(b *w[2]);       m[1] = (a *w[1])+(b *w[3]);
			m[2] = (c *w[0])+(d *w[2]);       m[3] = (c *w[1])+(d *w[3]);
			m[4] = (tx*w[0])+(ty*w[2])+w[4];  m[5] = (tx*w[1])+(ty*w[3])+w[5]; 
		},
		invert   : function(m:any    ) {  
			var a=m[0],b=m[1],c=m[2],d=m[3],tx=m[4],ty=m[5], adbc=a*d-b*c;
			m[0] = d/adbc;  m[1] = -b/adbc;  m[2] =-c/adbc;  m[3] =  a/adbc;
			m[4] = (c*ty - d*tx)/adbc;  m[5] = (b*tx - a*ty)/adbc;
		},
		multPoint: function(m:any, p:any ) {  var x=p[0],y=p[1];  return [x*m[0]+y*m[2]+m[4],   x*m[1]+y*m[3]+m[5]];  },
		multArray: function(m:any, a:any ) {  for(var i=0; i<a.length; i+=2) {  var x=a[i],y=a[i+1];  a[i]=x*m[0]+y*m[2]+m[4];  a[i+1]=x*m[1]+y*m[3]+m[5];  }  }
	}
	UDOC.C = {
		srgbGamma : function(x:any) {  return x < 0.0031308 ? 12.92 * x : 1.055 * Math.pow(x, 1.0 / 2.4) - 0.055;  },
		cmykToRgb : function(clr:any) { 
			var c=clr[0], m=clr[1], y=clr[2], k=clr[3];
			// return [1-Math.min(1,c+k), 1-Math.min(1, m+k), 1-Math.min(1,y+k)];
			var r = 255
			+ c * (-4.387332384609988  * c + 54.48615194189176  * m +  18.82290502165302  * y + 212.25662451639585 * k +  -285.2331026137004) 
			+ m * ( 1.7149763477362134 * m - 5.6096736904047315 * y + -17.873870861415444 * k - 5.497006427196366) 
			+ y * (-2.5217340131683033 * y - 21.248923337353073 * k +  17.5119270841813) 
			+ k * (-21.86122147463605  * k - 189.48180835922747);
			var g = 255
			+ c * (8.841041422036149   * c + 60.118027045597366 * m +  6.871425592049007  * y + 31.159100130055922 * k +  -79.2970844816548) 
			+ m * (-15.310361306967817 * m + 17.575251261109482 * y +  131.35250912493976 * k - 190.9453302588951) 
			+ y * (4.444339102852739   * y + 9.8632861493405    * k -  24.86741582555878) 
			+ k * (-20.737325471181034 * k - 187.80453709719578);
			var b = 255
			+ c * (0.8842522430003296  * c + 8.078677503112928  * m +  30.89978309703729  * y - 0.23883238689178934 * k + -14.183576799673286) 
			+ m * (10.49593273432072   * m + 63.02378494754052  * y +  50.606957656360734 * k - 112.23884253719248) 
			+ y * (0.03296041114873217 * y + 115.60384449646641 * k + -193.58209356861505)
			+ k * (-22.33816807309886  * k - 180.12613974708367);

			return [Math.max(0, Math.min(1, r/255)), Math.max(0, Math.min(1, g/255)), Math.max(0, Math.min(1, b/255))];
			//var iK = 1-c[3];  
			//return [(1-c[0])*iK, (1-c[1])*iK, (1-c[2])*iK];  
		},
		labToRgb  : function(lab:any) {
			var k = 903.3, e = 0.008856, L = lab[0], a = lab[1], b = lab[2];
			var fy = (L+16)/116, fy3 = fy*fy*fy;
			var fz = fy - b/200, fz3 = fz*fz*fz;
			var fx = a/500 + fy, fx3 = fx*fx*fx;
			var zr = fz3>e ? fz3 : (116*fz-16)/k;
			var yr = fy3>e ? fy3 : (116*fy-16)/k;
			var xr = fx3>e ? fx3 : (116*fx-16)/k;
				
			var X = xr*96.72, Y = yr*100, Z = zr*81.427, xyz = [X/100,Y/100,Z/100];
			var x2s = [3.1338561, -1.6168667, -0.4906146, -0.9787684,  1.9161415,  0.0334540, 0.0719453, -0.2289914,  1.4052427];
			
			var rgb = [ x2s[0]*xyz[0] + x2s[1]*xyz[1] + x2s[2]*xyz[2],
						x2s[3]*xyz[0] + x2s[4]*xyz[1] + x2s[5]*xyz[2],
						x2s[6]*xyz[0] + x2s[7]*xyz[1] + x2s[8]*xyz[2]  ];
			for(var i=0; i<3; i++) rgb[i] = Math.max(0, Math.min(1, UDOC.C.srgbGamma(rgb[i])));
			return rgb;
		}
	}
	
	UDOC.getState = function(crds:any):any {
		return {
			font : UDOC.getFont(),
			dd: {flat:1},  // device-dependent
			space :"/DeviceGray",
			// fill
			ca: 1,
			colr  : [0,0,0],
			sspace:"/DeviceGray",
			// stroke
			CA: 1,
			COLR : [0,0,0],
			bmode: "/Normal",
			SA:false, OPM:0, AIS:false, OP:false, op:false, SMask:"/None",
			lwidth : 1,
			lcap: 0,
			ljoin: 0,
			mlimit: 10,
			SM : 0.1,
			doff: 0,
			dash: [],
			ctm : [1,0,0,1,0,0],
			cpos: [0,0],
			pth : {cmds:[],crds:[]}, 
			cpth: crds ? UDOC.G.rectToPath(crds) : null  // clipping path
		};
	}
	
	UDOC.getFont = function() {
		return {
			Tc: 0, // character spacing
			Tw: 0, // word spacing
			Th:100, // horizontal scale
			Tl: 0, // leading
			Tf:"Helvetica-Bold", 
			Tfs:1, // font size
			Tmode:0, // rendering mode
			Trise:0, // rise
			Tk: 0,  // knockout
			Tal:0,  // align, 0: left, 1: right, 2: center
			Tun:0,  // 0: no, 1: underline
			
			Tm :[1,0,0,1,0,0],
			Tlm:[1,0,0,1,0,0],
			Trm:[1,0,0,1,0,0]
		};
	}


export let FromEMF:any = function()
{
}

FromEMF.Parse = function(buff:any, genv:any)
{
    buff = new Uint8Array(buff);  var off=0;
    //console.log(buff.slice(0,32));
    var prms:any = {fill:false, strk:false, bb:[0,0,1,1], wbb:[0,0,1,1], fnt:{nam:"Arial",hgh:25,und:false,orn:0}, tclr:[0,0,0], talg:0}, gst, tab = [], sts=[];
    
    var rI = FromEMF.B.readShort, rU = FromEMF.B.readUshort, rI32 = FromEMF.B.readInt, rU32 = FromEMF.B.readUint, rF32 = FromEMF.B.readFloat;	
    
    var opn=0;
    while(true) {
        var fnc = rU32(buff, off);  off+=4;
        var fnm = FromEMF.K[fnc]; 
        var siz = rU32(buff, off);  off+=4;
        
        //if(gst && isNaN(gst.ctm[0])) throw "e";
        //console.log(fnc,fnm,siz);
        
        var loff = off;
        
        //if(opn++==253) break;
        var obj:any = null, oid = 0;
        //console.log(fnm, siz);
        
        if(false) {}
        else if(fnm=="EOF") {  break;  }
        else if(fnm=="HEADER") {
            prms.bb = FromEMF._readBox(buff,loff);   loff+=16;  //console.log(fnm, prms.bb);
            genv.StartPage(prms.bb[0],prms.bb[1],prms.bb[2],prms.bb[3]);
            gst = UDOC.getState(prms.bb);	
        }
        else if(fnm=="SAVEDC") sts.push(JSON.stringify(gst), JSON.stringify(prms));
        else if(fnm=="RESTOREDC") {
            var dif = rI32(buff, loff);  loff+=4;
            while(dif<-1) {  sts.pop();  sts.pop();  }
            prms = JSON.parse(sts.pop());  gst = JSON.parse(sts.pop());
        }
        else if(fnm=="SELECTCLIPPATH") {  gst.cpth = JSON.parse(JSON.stringify(gst.pth));  }
        else if(["SETMAPMODE","SETPOLYFILLMODE","SETBKMODE"/*,"SETVIEWPORTEXTEX"*/,"SETICMMODE","SETROP2","EXTSELECTCLIPRGN"].indexOf(fnm)!=-1) {}
        //else if(fnm=="INTERSECTCLIPRECT") {  var r=prms.crct=FromEMF._readBox(buff, loff);  /*var y0=r[1],y1=r[3]; if(y0>y1){r[1]=y1; r[3]=y0;}*/ console.log(prms.crct);  }
        else if(fnm=="SETMITERLIMIT") gst.mlimit = rU32(buff, loff);
        else if(fnm=="SETTEXTCOLOR") prms.tclr = [buff[loff]/255, buff[loff+1]/255, buff[loff+2]/255]; 
        else if(fnm=="SETTEXTALIGN") prms.talg = rU32(buff, loff);
        else if(fnm=="SETVIEWPORTEXTEX" || fnm=="SETVIEWPORTORGEX") {
            if(prms.vbb==null) prms.vbb=[];
            var coff = fnm=="SETVIEWPORTORGEX" ? 0 : 2;
            prms.vbb[coff  ] = rI32(buff, loff);  loff+=4;
            prms.vbb[coff+1] = rI32(buff, loff);  loff+=4;
            //console.log(prms.vbb);
            if(fnm=="SETVIEWPORTEXTEX") FromEMF._updateCtm(prms, gst);
        }
        else if(fnm=="SETWINDOWEXTEX" || fnm=="SETWINDOWORGEX") {
            var coff = fnm=="SETWINDOWORGEX" ? 0 : 2;
            prms.wbb[coff  ] = rI32(buff, loff);  loff+=4;
            prms.wbb[coff+1] = rI32(buff, loff);  loff+=4;
            if(fnm=="SETWINDOWEXTEX") FromEMF._updateCtm(prms, gst);
        }
        //else if(fnm=="SETMETARGN") {}
        else if(fnm=="COMMENT") {  var ds = rU32(buff, loff);  loff+=4;  }
        
        else if(fnm=="SELECTOBJECT") {
            var ind = rU32(buff, loff);  loff+=4;
            //console.log(ind.toString(16), tab, tab[ind]);
            if     (ind==0x80000000) {  prms.fill=true ;  gst.colr=[1,1,1];  } // white brush
            else if(ind==0x80000005) {  prms.fill=false;  } // null brush
            else if(ind==0x80000007) {  prms.strk=true ;  prms.lwidth=1;  gst.COLR=[0,0,0];  } // black pen
            else if(ind==0x80000008) {  prms.strk=false;  } // null  pen
            else if(ind==0x8000000d) {} // system font
            else if(ind==0x8000000e) {}  // device default font
            else {
                var co:any = tab[ind];  //console.log(ind, co);
                if(co.t=="b") {
                    prms.fill=co.stl!=1;
                    if     (co.stl==0) {}
                    else if(co.stl==1) {}
                    else throw co.stl+" e";
                    gst.colr=co.clr;
                }
                else if(co.t=="p") {
                    prms.strk=co.stl!=5;
                    gst.lwidth = co.wid;
                    gst.COLR=co.clr;
                }
                else if(co.t=="f") {
                    prms.fnt = co;
                    gst.font.Tf = co.nam;
                    gst.font.Tfs = Math.abs(co.hgh);
                    gst.font.Tun = co.und;
                }
                else throw "e";
            }
        }
        else if(fnm=="DELETEOBJECT") {
            var ind = rU32(buff, loff);  loff+=4;
            if(tab[ind]!=null) tab[ind]=null;
            else throw "e";
        }
        else if(fnm=="CREATEBRUSHINDIRECT") {
            oid = rU32(buff, loff);  loff+=4;
            obj = {t:"b"};
            obj.stl = rU32(buff, loff);  loff+=4;
            obj.clr = [buff[loff]/255, buff[loff+1]/255, buff[loff+2]/255];  loff+=4;
            obj.htc = rU32(buff, loff);  loff+=4;
            //console.log(oid, obj);
        }
        else if(fnm=="CREATEPEN" || fnm=="EXTCREATEPEN") {
            oid = rU32(buff, loff);  loff+=4;
            obj = {t:"p"};
            if(fnm=="EXTCREATEPEN") {
                loff+=16;
                obj.stl = rU32(buff, loff);  loff+=4;
                obj.wid = rU32(buff, loff);  loff+=4;
                //obj.stl = rU32(buff, loff);  
                loff+=4;
            } else {
                obj.stl = rU32(buff, loff);  loff+=4;
                obj.wid = rU32(buff, loff);  loff+=4;  loff+=4;
            }
            obj.clr = [buff[loff]/255, buff[loff+1]/255, buff[loff+2]/255];  loff+=4;
        }
        else if(fnm=="EXTCREATEFONTINDIRECTW") {
            oid = rU32(buff, loff);  loff+=4;
            obj = {t:"f", nam:""};
            obj.hgh = rI32(buff, loff);  loff += 4;
            loff += 4*2;
            obj.orn = rI32(buff, loff)/10;  loff+=4;
            var wgh = rU32(buff, loff);  loff+=4;  //console.log(fnm, obj.orn, wgh);
            //console.log(rU32(buff,loff), rU32(buff,loff+4), buff.slice(loff,loff+8));
            obj.und = buff[loff+1];  obj.stk = buff[loff+2];  loff += 4*2;
            while(rU(buff,loff)!=0) {  obj.nam+=String.fromCharCode(rU(buff,loff));  loff+=2;  }
            if(wgh>500) obj.nam+="-Bold";
            //console.log(wgh, obj.nam);
        }
        else if(fnm=="EXTTEXTOUTW") {
            //console.log(buff.slice(loff-8, loff-8+siz));
            loff+=16;
            var mod = rU32(buff, loff);  loff+=4;  //console.log(mod);
            var scx = rF32(buff, loff);  loff+=4;
            var scy = rF32(buff, loff);  loff+=4;
            var rfx = rI32(buff, loff);  loff+=4;
            var rfy = rI32(buff, loff);  loff+=4;
            //console.log(mod, scx, scy,rfx,rfy);
            
            gst.font.Tm = [1,0,0,-1,0,0];
            UDOC.M.rotate(gst.font.Tm, prms.fnt.orn*Math.PI/180);
            UDOC.M.translate(gst.font.Tm, rfx, rfy);
            
            var alg = prms.talg;  //console.log(alg.toString(2));
            if     ((alg&6)==6) gst.font.Tal = 2;
            else if((alg&7)==0) gst.font.Tal = 0;
            else throw alg+" e";
            if((alg&24)==24) {}  // baseline
            else if((alg&24)==0) UDOC.M.translate(gst.font.Tm, 0, gst.font.Tfs);
            else throw "e";
            
            
            var crs = rU32(buff, loff);  loff+=4;
            var ofs = rU32(buff, loff);  loff+=4;
            var ops = rU32(buff, loff);  loff+=4;  //if(ops!=0) throw "e";
            //console.log(ofs,ops,crs);
            loff+=16;
            var ofD = rU32(buff, loff);  loff+=4;  //console.log(ops, ofD, loff, ofs+off-8);
            ofs += off-8;  //console.log(crs, ops);
            var str = "";
            for(var i=0; i<crs; i++) {  var cc=rU(buff,ofs+i*2);  str+=String.fromCharCode(cc);  };
            var oclr = gst.colr;  gst.colr = prms.tclr;
            //console.log(str, gst.colr, gst.font.Tm);
            //var otfs = gst.font.Tfs;  gst.font.Tfs *= 1/gst.ctm[0];
            genv.PutText(gst, str, str.length*gst.font.Tfs*0.5);  gst.colr=oclr;
            //gst.font.Tfs = otfs;
            //console.log(rfx, rfy, scx, ops, rcX, rcY, rcW, rcH, offDx, str);
        }
        else if(fnm=="BEGINPATH") {  UDOC.G.newPath(gst);  }
        else if(fnm=="ENDPATH"  ) {    }
        else if(fnm=="CLOSEFIGURE") UDOC.G.closePath(gst);
        else if(fnm=="MOVETOEX" ) {  UDOC.G.moveTo(gst, rI32(buff,loff), rI32(buff,loff+4));  }
        else if(fnm=="LINETO"   ) {  
            if(gst.pth.cmds.length==0) {  var im=gst.ctm.slice(0);  UDOC.M.invert(im);  var p = UDOC.M.multPoint(im, gst.cpos);  UDOC.G.moveTo(gst, p[0], p[1]);  }  
            UDOC.G.lineTo(gst, rI32(buff,loff), rI32(buff,loff+4));  }
        else if(fnm=="POLYGON" || fnm=="POLYGON16" || fnm=="POLYLINE" || fnm=="POLYLINE16" || fnm=="POLYLINETO" || fnm=="POLYLINETO16") {
            loff+=16;
            var ndf = fnm.startsWith("POLYGON"), isTo = fnm.indexOf("TO")!=-1;
            var cnt = rU32(buff, loff);  loff+=4;
            if(!isTo) UDOC.G.newPath(gst);
            loff = FromEMF._drawPoly(buff,loff,cnt,gst, fnm.endsWith("16")?2:4,  ndf, isTo);
            if(!isTo) FromEMF._draw(genv,gst,prms, ndf);
            //console.log(prms, gst.lwidth);
            //console.log(JSON.parse(JSON.stringify(gst.pth)));
        }
        else if(fnm=="POLYPOLYGON16") {
            loff+=16;
            var ndf = fnm.startsWith("POLYPOLYGON"), isTo = fnm.indexOf("TO")!=-1;
            var nop = rU32(buff, loff);  loff+=4;  loff+=4;
            var pi = loff;  loff+= nop*4;
            
            if(!isTo) UDOC.G.newPath(gst);
            for(var i=0; i<nop; i++) {
                var ppp = rU(buff, pi+i*4);
                loff = FromEMF._drawPoly(buff,loff,ppp,gst, fnm.endsWith("16")?2:4, ndf, isTo);
            }
            if(!isTo) FromEMF._draw(genv,gst,prms, ndf);
        }
        else if(fnm=="POLYBEZIER" || fnm=="POLYBEZIER16" || fnm=="POLYBEZIERTO" || fnm=="POLYBEZIERTO16") {
            loff+=16;
            var is16 = fnm.endsWith("16"), rC = is16?rI:rI32, nl = is16?2:4;
            var cnt = rU32(buff, loff);  loff+=4;
            if(fnm.indexOf("TO")==-1) {
                UDOC.G.moveTo(gst, rC(buff,loff), rC(buff,loff+nl));  loff+=2*nl;  cnt--;
            }
            while(cnt>0) {
                UDOC.G.curveTo(gst, rC(buff,loff), rC(buff,loff+nl), rC(buff,loff+2*nl), rC(buff,loff+3*nl), rC(buff,loff+4*nl), rC(buff,loff+5*nl) );
                loff+=6*nl;
                cnt-=3;
            }
            //console.log(JSON.parse(JSON.stringify(gst.pth)));
        }
        else if(fnm=="RECTANGLE" || fnm=="ELLIPSE") {
            UDOC.G.newPath(gst);
            var bx = FromEMF._readBox(buff, loff);
            if(fnm=="RECTANGLE") {
                UDOC.G.moveTo(gst, bx[0],bx[1]);
                UDOC.G.lineTo(gst, bx[2],bx[1]);
                UDOC.G.lineTo(gst, bx[2],bx[3]);
                UDOC.G.lineTo(gst, bx[0],bx[3]);
            }
            else {
                var x = (bx[0]+bx[2])/2, y = (bx[1]+bx[3])/2;
                UDOC.G.arc(gst,x,y,(bx[2]-bx[0])/2,0,2*Math.PI, false);
            }
            UDOC.G.closePath(gst);
            FromEMF._draw(genv,gst,prms, true);
            //console.log(prms, gst.lwidth);
        }
        else if(fnm=="FILLPATH"  ) genv.Fill(gst, false);
        else if(fnm=="STROKEPATH") genv.Stroke(gst);
        else if(fnm=="STROKEANDFILLPATH") {  genv.Fill(gst, false);  genv.Stroke(gst);  }
        else if(fnm=="SETWORLDTRANSFORM" || fnm=="MODIFYWORLDTRANSFORM") {
            var mat = [];
            for(var i=0; i<6; i++) mat.push(rF32(buff,loff+i*4));  loff+=24;
            //console.log(fnm, gst.ctm.slice(0), mat);
            if(fnm=="SETWORLDTRANSFORM") gst.ctm=mat;
            else {
                var mod = rU32(buff,loff);  loff+=4;
                if(mod==2) {  var om=gst.ctm;  gst.ctm=mat;  UDOC.M.concat(gst.ctm, om);  }
                else throw "e";
            }
        }
        else if(fnm=="SETSTRETCHBLTMODE") {  var sm = rU32(buff, loff);  loff+=4;  }
        else if(fnm=="STRETCHDIBITS") {
            var bx = FromEMF._readBox(buff, loff);  loff+=16;
            var xD = rI32(buff, loff);  loff+=4;
            var yD = rI32(buff, loff);  loff+=4;
            var xS = rI32(buff, loff);  loff+=4;
            var yS = rI32(buff, loff);  loff+=4;
            var wS = rI32(buff, loff);  loff+=4;
            var hS = rI32(buff, loff);  loff+=4;
            var ofH = rU32(buff, loff)+off-8;  loff+=4;
            var szH = rU32(buff, loff);  loff+=4;
            var ofB = rU32(buff, loff)+off-8;  loff+=4;
            var szB = rU32(buff, loff);  loff+=4;
            var usg = rU32(buff, loff);  loff+=4;  if(usg!=0) throw "e";
            var bop = rU32(buff, loff);  loff+=4;
            var wD = rI32(buff, loff);  loff+=4;
            var hD = rI32(buff, loff);  loff+=4;  //console.log(bop, wD, hD);
            
            //console.log(ofH, szH, ofB, szB, ofH+40);
            //console.log(bx, xD,yD,wD,hD);
            //console.log(xS,yS,wS,hS);
            //console.log(ofH,szH,ofB,szB,usg,bop);
            
            var hl = rU32(buff, ofH);  ofH+=4;
            var w  = rU32(buff, ofH);  ofH+=4;
            var h  = rU32(buff, ofH);  ofH+=4;  if(w!=wS || h!=hS) throw "e";
            var ps = rU  (buff, ofH);  ofH+=2;
            var bc = rU  (buff, ofH);  ofH+=2;  if(bc!=8 && bc!=24 && bc!=32) throw bc+" e";
            var cpr= rU32(buff, ofH);  ofH+=4;  if(cpr!=0) throw cpr+" e";
            var sz = rU32(buff, ofH);  ofH+=4;
            var xpm= rU32(buff, ofH);  ofH+=4;
            var ypm= rU32(buff, ofH);  ofH+=4;
            var cu = rU32(buff, ofH);  ofH+=4;
            var ci = rU32(buff, ofH);  ofH+=4;  //console.log(hl, w, h, ps, bc, cpr, sz, xpm, ypm, cu, ci);
            
            //console.log(hl,w,h,",",xS,yS,wS,hS,",",xD,yD,wD,hD,",",xpm,ypm);
            
            var rl = Math.floor(((w * ps * bc + 31) & ~31) / 8);
            var img = new Uint8Array(w*h*4);
            if(bc==8) {
                for(var y=0; y<h; y++) 
                    for(var x=0; x<w; x++) {
                        var qi = (y*w+x)<<2, ind:any = buff[ofB+(h-1-y)*rl+x]<<2;
                        img[qi  ] = buff[ofH+ind+2];
                        img[qi+1] = buff[ofH+ind+1];
                        img[qi+2] = buff[ofH+ind+0];
                        img[qi+3] = 255;
                    }
            }
            if(bc==24) {
                for(var y=0; y<h; y++) 
                    for(var x=0; x<w; x++) {
                        var qi = (y*w+x)<<2, ti=ofB+(h-1-y)*rl+x*3;
                        img[qi  ] = buff[ti+2];
                        img[qi+1] = buff[ti+1];
                        img[qi+2] = buff[ti+0];
                        img[qi+3] = 255;
                    }
            }
            if(bc==32) {
                for(var y=0; y<h; y++) 
                    for(var x=0; x<w; x++) {
                        var qi = (y*w+x)<<2, ti=ofB+(h-1-y)*rl+x*4;
                        img[qi  ] = buff[ti+2];
                        img[qi+1] = buff[ti+1];
                        img[qi+2] = buff[ti+0];
                        img[qi+3] = buff[ti+3];
                    }
            }
            
            var ctm = gst.ctm.slice(0);
            gst.ctm = [1,0,0,1,0,0];
            UDOC.M.scale(gst.ctm, wD, -hD);
            UDOC.M.translate(gst.ctm, xD, yD+hD);
            UDOC.M.concat(gst.ctm, ctm);
            genv.PutImage(gst, img, w, h);
            gst.ctm = ctm;
        }
        else {
            console.log(fnm, siz);
        }
        
        if(obj!=null) tab[oid]=obj;
        
        off+=siz-8;
    }
    //genv.Stroke(gst);
    genv.ShowPage();  genv.Done();
}
FromEMF._readBox = function(buff:any, off:any) {  var b=[];  for(var i=0; i<4; i++) b[i] = FromEMF.B.readInt(buff,off+i*4);  return b;  }	

FromEMF._updateCtm = function(prms:any, gst:any) {
    var mat = [1,0,0,1,0,0];
    var wbb = prms.wbb, bb = prms.bb, vbb=(prms.vbb && prms.vbb.length==4) ? prms.vbb:prms.bb;
    
    //var y0 = bb[1], y1 = bb[3];  bb[1]=Math.min(y0,y1);  bb[3]=Math.max(y0,y1);
    
    UDOC.M.translate(mat, -wbb[0],-wbb[1]);
    UDOC.M.scale(mat, 1/wbb[2], 1/wbb[3]);
    
    UDOC.M.scale(mat, vbb[2], vbb[3]);
    //UDOC.M.scale(mat, vbb[2]/(bb[2]-bb[0]), vbb[3]/(bb[3]-bb[1]));
    
    //UDOC.M.scale(mat, bb[2]-bb[0],bb[3]-bb[1]);
    
    gst.ctm = mat;
}
FromEMF._draw = function(genv:any, gst:any, prms:any, needFill:any) {
    if(prms.fill && needFill     ) genv.Fill  (gst, false);
    if(prms.strk && gst.lwidth!=0) genv.Stroke(gst);
}
FromEMF._drawPoly = function(buff:any, off:any, ppp:any, gst:any, nl:any, clos:any, justLine:any) {
    var rS = nl==2 ? FromEMF.B.readShort : FromEMF.B.readInt;
    for(var j=0; j<ppp; j++) {
        var px = rS(buff, off);  off+=nl;  
        var py = rS(buff, off);  off+=nl;
        if(j==0 && !justLine) UDOC.G.moveTo(gst,px,py);  else UDOC.G.lineTo(gst,px,py);
    }
    if(clos) UDOC.G.closePath(gst);
    return off;
}

FromEMF.B = {
    uint8 : new Uint8Array(4),
    readShort  : function(buff:any,p:any):any  {  var u8=FromEMF.B.uint8;  u8[0]=buff[p];  u8[1]=buff[p+1];  return FromEMF.B.int16 [0];  },
    readUshort : function(buff:any,p:any):any  {  var u8=FromEMF.B.uint8;  u8[0]=buff[p];  u8[1]=buff[p+1];  return FromEMF.B.uint16[0];  },
    readInt    : function(buff:any,p:any):any  {  var u8=FromEMF.B.uint8;  u8[0]=buff[p];  u8[1]=buff[p+1];  u8[2]=buff[p+2];  u8[3]=buff[p+3];  return FromEMF.B.int32 [0];  },
    readUint   : function(buff:any,p:any):any  {  var u8=FromEMF.B.uint8;  u8[0]=buff[p];  u8[1]=buff[p+1];  u8[2]=buff[p+2];  u8[3]=buff[p+3];  return FromEMF.B.uint32[0];  },
    readFloat  : function(buff:any,p:any):any  {  var u8=FromEMF.B.uint8;  u8[0]=buff[p];  u8[1]=buff[p+1];  u8[2]=buff[p+2];  u8[3]=buff[p+3];  return FromEMF.B.flot32[0];  },
    readASCII  : function(buff:any,p:any,l:any):any {  var s = "";  for(var i=0; i<l; i++) s += String.fromCharCode(buff[p+i]);  return s;    }
}
FromEMF.B.int16  = new Int16Array (FromEMF.B.uint8.buffer);
FromEMF.B.uint16 = new Uint16Array(FromEMF.B.uint8.buffer);
FromEMF.B.int32  = new Int32Array (FromEMF.B.uint8.buffer);
FromEMF.B.uint32 = new Uint32Array(FromEMF.B.uint8.buffer);
FromEMF.B.flot32 = new Float32Array(FromEMF.B.uint8.buffer);


FromEMF.C = {
    EMR_HEADER : 0x00000001,
    EMR_POLYBEZIER : 0x00000002,
    EMR_POLYGON : 0x00000003,
    EMR_POLYLINE : 0x00000004,
    EMR_POLYBEZIERTO : 0x00000005,
    EMR_POLYLINETO : 0x00000006,
    EMR_POLYPOLYLINE : 0x00000007,
    EMR_POLYPOLYGON : 0x00000008,
    EMR_SETWINDOWEXTEX : 0x00000009,
    EMR_SETWINDOWORGEX : 0x0000000A,
    EMR_SETVIEWPORTEXTEX : 0x0000000B,
    EMR_SETVIEWPORTORGEX : 0x0000000C,
    EMR_SETBRUSHORGEX : 0x0000000D,
    EMR_EOF : 0x0000000E,
    EMR_SETPIXELV : 0x0000000F,
    EMR_SETMAPPERFLAGS : 0x00000010,
    EMR_SETMAPMODE : 0x00000011,
    EMR_SETBKMODE : 0x00000012,
    EMR_SETPOLYFILLMODE : 0x00000013,
    EMR_SETROP2 : 0x00000014,
    EMR_SETSTRETCHBLTMODE : 0x00000015,
    EMR_SETTEXTALIGN : 0x00000016,
    EMR_SETCOLORADJUSTMENT : 0x00000017,
    EMR_SETTEXTCOLOR : 0x00000018,
    EMR_SETBKCOLOR : 0x00000019,
    EMR_OFFSETCLIPRGN : 0x0000001A,
    EMR_MOVETOEX : 0x0000001B,
    EMR_SETMETARGN : 0x0000001C,
    EMR_EXCLUDECLIPRECT : 0x0000001D,
    EMR_INTERSECTCLIPRECT : 0x0000001E,
    EMR_SCALEVIEWPORTEXTEX : 0x0000001F,
    EMR_SCALEWINDOWEXTEX : 0x00000020,
    EMR_SAVEDC : 0x00000021,
    EMR_RESTOREDC : 0x00000022,
    EMR_SETWORLDTRANSFORM : 0x00000023,
    EMR_MODIFYWORLDTRANSFORM : 0x00000024,
    EMR_SELECTOBJECT : 0x00000025,
    EMR_CREATEPEN : 0x00000026,
    EMR_CREATEBRUSHINDIRECT : 0x00000027,
    EMR_DELETEOBJECT : 0x00000028,
    EMR_ANGLEARC : 0x00000029,
    EMR_ELLIPSE : 0x0000002A,
    EMR_RECTANGLE : 0x0000002B,
    EMR_ROUNDRECT : 0x0000002C,
    EMR_ARC : 0x0000002D,
    EMR_CHORD : 0x0000002E,
    EMR_PIE : 0x0000002F,
    EMR_SELECTPALETTE : 0x00000030,
    EMR_CREATEPALETTE : 0x00000031,
    EMR_SETPALETTEENTRIES : 0x00000032,
    EMR_RESIZEPALETTE : 0x00000033,
    EMR_REALIZEPALETTE : 0x00000034,
    EMR_EXTFLOODFILL : 0x00000035,
    EMR_LINETO : 0x00000036,
    EMR_ARCTO : 0x00000037,
    EMR_POLYDRAW : 0x00000038,
    EMR_SETARCDIRECTION : 0x00000039,
    EMR_SETMITERLIMIT : 0x0000003A,
    EMR_BEGINPATH : 0x0000003B,
    EMR_ENDPATH : 0x0000003C,
    EMR_CLOSEFIGURE : 0x0000003D,
    EMR_FILLPATH : 0x0000003E,
    EMR_STROKEANDFILLPATH : 0x0000003F,
    EMR_STROKEPATH : 0x00000040,
    EMR_FLATTENPATH : 0x00000041,
    EMR_WIDENPATH : 0x00000042,
    EMR_SELECTCLIPPATH : 0x00000043,
    EMR_ABORTPATH : 0x00000044,
    EMR_COMMENT : 0x00000046,
    EMR_FILLRGN : 0x00000047,
    EMR_FRAMERGN : 0x00000048,
    EMR_INVERTRGN : 0x00000049,
    EMR_PAINTRGN : 0x0000004A,
    EMR_EXTSELECTCLIPRGN : 0x0000004B,
    EMR_BITBLT : 0x0000004C,
    EMR_STRETCHBLT : 0x0000004D,
    EMR_MASKBLT : 0x0000004E,
    EMR_PLGBLT : 0x0000004F,
    EMR_SETDIBITSTODEVICE : 0x00000050,
    EMR_STRETCHDIBITS : 0x00000051,
    EMR_EXTCREATEFONTINDIRECTW : 0x00000052,
    EMR_EXTTEXTOUTA : 0x00000053,
    EMR_EXTTEXTOUTW : 0x00000054,
    EMR_POLYBEZIER16 : 0x00000055,
    EMR_POLYGON16 : 0x00000056,
    EMR_POLYLINE16 : 0x00000057,
    EMR_POLYBEZIERTO16 : 0x00000058,
    EMR_POLYLINETO16 : 0x00000059,
    EMR_POLYPOLYLINE16 : 0x0000005A,
    EMR_POLYPOLYGON16 : 0x0000005B,
    EMR_POLYDRAW16 : 0x0000005C,
    EMR_CREATEMONOBRUSH : 0x0000005D,
    EMR_CREATEDIBPATTERNBRUSHPT : 0x0000005E,
    EMR_EXTCREATEPEN : 0x0000005F,
    EMR_POLYTEXTOUTA : 0x00000060,
    EMR_POLYTEXTOUTW : 0x00000061,
    EMR_SETICMMODE : 0x00000062,
    EMR_CREATECOLORSPACE : 0x00000063,
    EMR_SETCOLORSPACE : 0x00000064,
    EMR_DELETECOLORSPACE : 0x00000065,
    EMR_GLSRECORD : 0x00000066,
    EMR_GLSBOUNDEDRECORD : 0x00000067,
    EMR_PIXELFORMAT : 0x00000068,
    EMR_DRAWESCAPE : 0x00000069,
    EMR_EXTESCAPE : 0x0000006A,
    EMR_SMALLTEXTOUT : 0x0000006C,
    EMR_FORCEUFIMAPPING : 0x0000006D,
    EMR_NAMEDESCAPE : 0x0000006E,
    EMR_COLORCORRECTPALETTE : 0x0000006F,
    EMR_SETICMPROFILEA : 0x00000070,
    EMR_SETICMPROFILEW : 0x00000071,
    EMR_ALPHABLEND : 0x00000072,
    EMR_SETLAYOUT : 0x00000073,
    EMR_TRANSPARENTBLT : 0x00000074,
    EMR_GRADIENTFILL : 0x00000076,
    EMR_SETLINKEDUFIS : 0x00000077,
    EMR_SETTEXTJUSTIFICATION : 0x00000078,
    EMR_COLORMATCHTOTARGETW : 0x00000079,
    EMR_CREATECOLORSPACEW : 0x0000007A
};
FromEMF.K = [];

// (function() {
//     var inp, out, stt;
//     inp = FromEMF.C;   out = FromEMF.K;   stt=4;
//     for(var p in inp) out[inp[p]] = p.slice(stt);
// }  )();



export let ToContext2D:any = function (needPage:any, scale:any)
{
    this.canvas = document.createElement("canvas");
    this.ctx = this.canvas.getContext("2d");
    this.bb = null;
    this.currPage = 0;
    this.needPage = needPage;
    this.scale = scale;
}
ToContext2D.prototype.StartPage = function(x:any,y:any,w:any,h:any) {
    if(this.currPage!=this.needPage) return;
    this.bb = [x,y,w,h];
    var scl = this.scale, dpr = window.devicePixelRatio;
    var cnv = this.canvas, ctx = this.ctx;
    cnv.width = Math.round(w*scl);  cnv.height = Math.round(h*scl);
    ctx.translate(0,h*scl);  ctx.scale(scl,-scl);
    cnv.setAttribute("style", "border:1px solid; width:"+(cnv.width/dpr)+"px; height:"+(cnv.height/dpr)+"px");
}
ToContext2D.prototype.Fill = function(gst:any, evenOdd:any) {
    if(this.currPage!=this.needPage) return;
    var ctx = this.ctx;
    ctx.beginPath();
    this._setStyle(gst, ctx);
    this._draw(gst.pth, ctx);
    ctx.fill();
}
ToContext2D.prototype.Stroke = function(gst:any) {
    if(this.currPage!=this.needPage) return;
    var ctx = this.ctx;
    ctx.beginPath();
    this._setStyle(gst, ctx);
    this._draw(gst.pth, ctx);
    ctx.stroke();
}
ToContext2D.prototype.PutText = function(gst:any, str:any, stw:any) {
    if(this.currPage!=this.needPage) return;
    var scl = this._scale(gst.ctm);
    var ctx = this.ctx;
    this._setStyle(gst, ctx);
    ctx.save();
    var m = [1,0,0,-1,0,0];  this._concat(m, gst.font.Tm);  this._concat(m, gst.ctm);
    //console.log(str, m, gst);  throw "e";
    ctx.transform(m[0],m[1],m[2],m[3],m[4],m[5]);
    ctx.fillText(str,0,0);
    ctx.restore();
}
ToContext2D.prototype.PutImage = function(gst:any, buff:any, w:any, h:any, msk:any) {
    if(this.currPage!=this.needPage) return;
    var ctx = this.ctx;
    
    if(buff.length==w*h*4) {
        buff = buff.slice(0);
        if(msk && msk.length==w*h*4) for(var i=0; i<buff.length; i+=4) buff[i+3] = msk[i+1];
        
        var cnv = document.createElement("canvas"), cctx = cnv.getContext("2d");
        cnv.width = w;  cnv.height = h;
        var imgd = cctx.createImageData(w,h);
        for(var i=0; i<buff.length; i++) imgd.data[i]=buff[i];
        cctx.putImageData(imgd,0,0);
        
        ctx.save();
        var m = [1,0,0,1,0,0];  this._concat(m, [1/w,0,0,-1/h,0,1]);  this._concat(m, gst.ctm);
        ctx.transform(m[0],m[1],m[2],m[3],m[4],m[5]);
        ctx.drawImage(cnv,0,0);
        ctx.restore();
    }
}
ToContext2D.prototype.ShowPage = function() {  this.currPage++;  }
ToContext2D.prototype.Done = function() {}


function _flt(n:any)  {  return ""+parseFloat(n.toFixed(2));  }

ToContext2D.prototype._setStyle = function(gst:any, ctx:any) {
    var scl = this._scale(gst.ctm);
    ctx.fillStyle = this._getFill(gst.colr, gst.ca, ctx);
    ctx.strokeStyle=this._getFill(gst.COLR, gst.CA, ctx);
    
    ctx.lineCap = ["butt","round","square"][gst.lcap];
    ctx.lineJoin= ["miter","round","bevel"][gst.ljoin];
    ctx.lineWidth=gst.lwidth*scl;
    var dsh = gst.dash.slice(0);  for(var i=0; i<dsh.length; i++) dsh[i] = _flt(dsh[i]*scl);
    ctx.setLineDash(dsh); 
    ctx.miterLimit = gst.mlimit*scl;
    
    var fn = gst.font.Tf, ln = fn.toLowerCase();
    var p0 = ln.indexOf("bold")!=-1 ? "bold " : "";
    var p1 = (ln.indexOf("italic")!=-1 || ln.indexOf("oblique")!=-1) ? "italic " : "";
    ctx.font = p0+p1 + gst.font.Tfs+"px \""+fn+"\"";
}
ToContext2D.prototype._getFill = function(colr:any, ca:any, ctx:any)
{
    if(colr.typ==null) return this._colr(colr,ca);
    else {
        var grd = colr, crd = grd.crds, mat = grd.mat, scl=this._scale(mat), gf;
        if     (grd.typ=="lin") {
            var p0 = this._multPoint(mat,crd.slice(0,2)), p1 = this._multPoint(mat,crd.slice(2));
            gf=ctx.createLinearGradient(p0[0],p0[1],p1[0],p1[1]);
        }
        else if(grd.typ=="rad") {
            var p0 = this._multPoint(mat,crd.slice(0,2)), p1 = this._multPoint(mat,crd.slice(3));
            gf=ctx.createRadialGradient(p0[0],p0[1],crd[2]*scl,p1[0],p1[1],crd[5]*scl);
        }
        for(var i=0; i<grd.grad.length; i++)  gf.addColorStop(grd.grad[i][0],this._colr(grd.grad[i][1], ca));
        return gf;
    }
}
ToContext2D.prototype._colr  = function(c:any,a:any) {  return "rgba("+Math.round(c[0]*255)+","+Math.round(c[1]*255)+","+Math.round(c[2]*255)+","+a+")";  };
ToContext2D.prototype._scale = function(m:any)  {  return Math.sqrt(Math.abs(m[0]*m[3]-m[1]*m[2]));  };
ToContext2D.prototype._concat= function(m:any,w:any  ) {  
        var a=m[0],b=m[1],c=m[2],d=m[3],tx=m[4],ty=m[5];
        m[0] = (a *w[0])+(b *w[2]);       m[1] = (a *w[1])+(b *w[3]);
        m[2] = (c *w[0])+(d *w[2]);       m[3] = (c *w[1])+(d *w[3]);
        m[4] = (tx*w[0])+(ty*w[2])+w[4];  m[5] = (tx*w[1])+(ty*w[3])+w[5]; 
}
ToContext2D.prototype._multPoint= function(m:any, p:any) {  var x=p[0],y=p[1];  return [x*m[0]+y*m[2]+m[4],   x*m[1]+y*m[3]+m[5]];  },
ToContext2D.prototype._draw  = function(path:any, ctx:any)
{
    var c = 0, crds = path.crds;
    for(var j=0; j<path.cmds.length; j++) {
        var cmd = path.cmds[j];
        if     (cmd=="M") {  ctx.moveTo(crds[c], crds[c+1]);  c+=2;  }
        else if(cmd=="L") {  ctx.lineTo(crds[c], crds[c+1]);  c+=2;  }
        else if(cmd=="C") {  ctx.bezierCurveTo(crds[c], crds[c+1], crds[c+2], crds[c+3], crds[c+4], crds[c+5]);  c+=6;  }
        else if(cmd=="Q") {  ctx.quadraticCurveTo(crds[c], crds[c+1], crds[c+2], crds[c+3]);  c+=4;  }
        else if(cmd=="Z") {  ctx.closePath();  }
    }
}
