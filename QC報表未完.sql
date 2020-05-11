Select tracking.timein,tracking.timeout, tracking.users, tracking.sn,value  from tracking 
left join paravalue on tracking.sn = paravalue.sn and parameter = 'QCstatus'
 where tracking.station = '0330' and tracking.timeout between '2020/04/03' and '2020/04/10' and  parameter = 'QCStatus' and value is not null