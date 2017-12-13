//This source is created by ExcelStateChartConverter.exe. Source : doc\ball_state_table.xlsx
'
Sub BallStateSelect(func As String, bFirst As Boolean)
    Select Case func
        Case "S_NONE"
        	S_NONE(bFirst)
        Case "S_IDLE"
        	S_IDLE(bFirst)
        Case "S_MOVE"
        	S_MOVE(bFirst)
        Case "S_HITWALL"
        	S_HITWALL(bFirst)
        Case "S_HITDIR"
        	S_HITDIR(bFirst)
        Case "S_MOVE_DIR"
        	S_MOVE_DIR(bFirst)
        Case "S_HITYAJI"
        	S_HITYAJI(bFirst)
        Case "S_HITTARGET"
        	S_HITTARGET(bFirst)
        Case "S_HITENEMY"
        	S_HITENEMY(bFirst)
        Case "S_HITBROCK"
        	S_HITBROCK(bFirst)
        Case "S_HITOBS"
        	S_HITOBS(bFirst)
        Case "S_VANISH"
        	S_VANISH(bFirst)
        
    EndSelect
EndSub
'
void S_NONE(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(null);
    }
    if (HasNextState())
    {
        GoNextState();
    }
}
void S_IDLE(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(null);
    }
    vec_reset(bFirst);
    pos_at_player(bFirst);
    br_input_vec(S_MOVE);
    timeout_reset(bFirst);
    if (HasNextState())
    {
        timeout_start();
        GoNextState();
    }
}
void S_MOVE(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(null);
    }
    hit_check();
    pos_move_if_nothit();
    br_if_wall(S_HITWALL);
    br_if_dirbr(S_HITDIR);
    br_if_yajibr(S_HITYAJI);
    br_if_target(S_HITTARGET);
    br_if_enemy(S_HITENEMY);
    br_if_brock(S_HITBROCK);
    br_if_obs(S_HITOBS);
    br_if_wall_block_obs_from_dir(S_VANISH);
    timeout(S_VANISH);
    if (HasNextState())
    {
        nowait();
        GoNextState();
    }
}
void S_HITWALL(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(S_MOVE);
    }
    vec_normal_reflect();
    pos_hit_point();
    timeout(S_VANISH);
    if (HasNextState())
    {
        GoNextState();
    }
}
void S_HITDIR(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(S_MOVE_DIR);
    }
    hit_dir_save();
    vec_dir_reflect();
    pos_hit_point();
    timeout(S_VANISH);
    if (HasNextState())
    {
        GoNextState();
    }
}
void S_MOVE_DIR(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(null);
    }
    pos_move_if_in_dir();
    br_if_not_in_dir(S_MOVE);
    timeout(S_VANISH);
    if (HasNextState())
    {
        nowait();
        hit_dir_reset();
        GoNextState();
    }
}
void S_HITYAJI(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(S_VANISH);
    }
    yaji_change();
    timeout(S_VANISH);
    if (HasNextState())
    {
        GoNextState();
    }
}
void S_HITTARGET(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(S_VANISH);
    }
    target_vanish();
    timeout(S_VANISH);
    if (HasNextState())
    {
        GoNextState();
    }
}
void S_HITENEMY(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(S_MOVE);
    }
    vec_normal_reflect();
    pos_hit_point();
    enemy_vanish();
    timeout(S_VANISH);
    if (HasNextState())
    {
        GoNextState();
    }
}
void S_HITBROCK(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(S_MOVE);
    }
    vec_normal_reflect();
    pos_hit_point();
    brock_vanish();
    timeout(S_VANISH);
    if (HasNextState())
    {
        GoNextState();
    }
}
void S_HITOBS(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(S_MOVE);
    }
    vec_normal_reflect();
    pos_hit_point();
    timeout(S_VANISH);
    if (HasNextState())
    {
        GoNextState();
    }
}
void S_VANISH(bool bFirst)
{
    if (bFirst)
    {
        SetNextState(S_NONE);
    }
    ball_vanish();
    if (HasNextState())
    {
        GoNextState();
    }
}
