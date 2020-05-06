SELECT    queue.name,
    initiationmethod,
    CASE WHEN agent is not null THEN 1 ELSE 0 END as Agente,
    replace(replace(queue.enqueuetimestamp,'T',' '),'Z','') as enqueuetimestamp,
    replace(replace(agent.connectedtoagenttimestamp,'T',' '),'Z','') as connectedtoagenttimestamp,
    replace(replace(disconnecttimestamp,'T',' '),'Z','') as disconnecttimestamp,
    queue.duration,
    agent.agentinteractionduration,
    agent.username
    FROM AwsDataCatalog.asd_amr_db.all_ctr_data
    WHERE substring(queue.name,1,5) = 'Femsa'
    And cast(replace(replace(queue.enqueuetimestamp,'T',' '),'Z','') as timestamp) between cast('2020-04-28 05:00:00' as timestamp) and cast('2020-04-29 04:59:59' as timestamp)
    order by cast(replace(replace(queue.enqueuetimestamp,'T',' '),'Z','') as timestamp)