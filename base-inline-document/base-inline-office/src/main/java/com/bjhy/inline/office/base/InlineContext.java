package com.bjhy.inline.office.base;

import java.net.InetAddress;
import java.net.NetworkInterface;
import java.net.UnknownHostException;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.Set;

import javax.management.AttributeNotFoundException;
import javax.management.InstanceNotFoundException;
import javax.management.MBeanException;
import javax.management.MBeanServer;
import javax.management.MBeanServerFactory;
import javax.management.ObjectName;
import javax.management.ReflectionException;
import javax.servlet.ServletRequestEvent;
import javax.servlet.ServletRequestListener;

import org.springframework.stereotype.Component;

/**
 * 得到contextPath
 * @author wubo
 *
 */
@Component
public class InlineContext implements ServletRequestListener{
	
	public static String contextPath; //当前系统的contextPath
	
	@Override
	public void requestInitialized(ServletRequestEvent servletRequestEvent) {
		if(contextPath == null){
			try {
				String host = getHostAddress();
//				String port = getServerPort(true);
				int port = servletRequestEvent.getServletRequest().getServerPort();
				String context = servletRequestEvent.getServletRequest().getServletContext().getContextPath();
				contextPath = "http://"+host+":"+port+context;
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
	
	@Override
	public void requestDestroyed(ServletRequestEvent arg0) {
		// TODO Auto-generated method stub
		
	}
	

	private String getHostAddress() throws UnknownHostException {  
        Enumeration<NetworkInterface> netInterfaces = null;  
        try {  
            netInterfaces = NetworkInterface.getNetworkInterfaces();  
            while (netInterfaces.hasMoreElements()) {  
                NetworkInterface ni = netInterfaces.nextElement();  
                Enumeration<InetAddress> ips = ni.getInetAddresses();  
                while (ips.hasMoreElements()) {  
                    InetAddress ip = ips.nextElement();  
                    if (ip.isSiteLocalAddress()) {  
                        return ip.getHostAddress();  
                    }  
                }  
            }  
        } catch (Exception e) {  
            e.printStackTrace();
        }  
        return InetAddress.getLocalHost().getHostAddress();  
    }  
	
	
	/** 
     * 获取服务端口号 
     * @return 端口号 
     * @throws ReflectionException 
     * @throws MBeanException 
     * @throws InstanceNotFoundException 
     * @throws AttributeNotFoundException 
     */  
    private String getServerPort(boolean secure) throws AttributeNotFoundException, InstanceNotFoundException, MBeanException, ReflectionException {  
        MBeanServer mBeanServer = null;  
        if (MBeanServerFactory.findMBeanServer(null).size() > 0) {  
            mBeanServer = (MBeanServer)MBeanServerFactory.findMBeanServer(null).get(0);  
        }  
          
        if (mBeanServer == null) {  
            System.out.println("调用findMBeanServer查询到的结果为null");
            return "";  
        }  
          
        Set<ObjectName> names = null;  
        try {  
            names = mBeanServer.queryNames(new ObjectName("Catalina:type=Connector,*"), null);  
        } catch (Exception e) {  
            return "";  
        }  
        Iterator<ObjectName> it = names.iterator();  
        ObjectName oname = null;  
        while (it.hasNext()) {  
            oname = (ObjectName)it.next();  
            String protocol = (String)mBeanServer.getAttribute(oname, "protocol");  
            String scheme = (String)mBeanServer.getAttribute(oname, "scheme");  
            Boolean secureValue = (Boolean)mBeanServer.getAttribute(oname, "secure");  
            Boolean SSLEnabled = (Boolean)mBeanServer.getAttribute(oname, "SSLEnabled");  
            if (SSLEnabled != null && SSLEnabled) {// tomcat6开始用SSLEnabled  
                secureValue = true;// SSLEnabled=true但secure未配置的情况  
                scheme = "https";  
            }  
            if (protocol != null && ("HTTP/1.1".equals(protocol) || protocol.contains("http"))) {  
                if (secure && "https".equals(scheme) && secureValue) {  
                    return ((Integer)mBeanServer.getAttribute(oname, "port")).toString();  
                } else if (!secure && !"https".equals(scheme) && !secureValue) {  
                    return ((Integer)mBeanServer.getAttribute(oname, "port")).toString();  
                }  
            }  
        }  
        return "";  
    }

}
