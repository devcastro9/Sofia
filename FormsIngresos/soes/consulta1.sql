select idFuncionario, paterno, materno, nombres, codigo_prisma
from tmp_rc_personal
where idFuncionario in (
select distinct dso_nro_veces
from detalle_soes
where dso_nro_veces <> 0 )
order by paterno, materno, nombres