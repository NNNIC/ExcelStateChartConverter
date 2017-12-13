//This source is created by ExcelStateChartConverter.exe. Source : doc\ball_state_table.xlsx
public enum BallState
{
    S_NONE
    ,S_IDLE
    ,S_MOVE
    ,S_HITWALL
    ,S_HITDIR
    ,S_MOVE_DIR
    ,S_HITYAJI
    ,S_HITTARGET
    ,S_HITENEMY
    ,S_HITBROCK
    ,S_HITOBS
    ,S_VANISH

}
public partial class BallStateControl {
    /*
      無指定
    */
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
    /*
      準備中
    */
    void S_IDLE(bool bFirst)
    {
        if (bFirst)
        {
            /*
              分岐で決定
            */
            SetNextState(null);
        }
        /*
          ベクトルリセット
        */
        vec_reset(bFirst);
        /*
          プレイヤ側に設置
        */
        pos_at_player(bFirst);
        /*
          ベクトル入力で分岐
        */
        br_input_vec(S_MOVE);
        timeout_reset(bFirst);
        if (HasNextState())
        {
            timeout_start();
            GoNextState();
        }
    }
    /*
      移動
    */
    void S_MOVE(bool bFirst)
    {
        if (bFirst)
        {
            /*
              分岐で決定
            */
            SetNextState(null);
        }
        /*
          進行方向の衝突確認
        */
        hit_check();
        /*
          衝突がなければ移動
        */
        pos_move_if_nothit();
        /*
          ヒット物で分岐
        */
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
    /*
      壁に衝突
    */
    void S_HITWALL(bool bFirst)
    {
        if (bFirst)
        {
            SetNextState(S_MOVE);
        }
        /*
          ノーマル反射
        */
        vec_normal_reflect();
        /*
          衝突ポイントへ移動
        */
        pos_hit_point();
        timeout(S_VANISH);
        if (HasNextState())
        {
            GoNextState();
        }
    }
    /*
      方向ブロックに衝突
    */
    void S_HITDIR(bool bFirst)
    {
        if (bFirst)
        {
            /*
              方向ブロックの中を進むと仮定
            */
            SetNextState(S_MOVE_DIR);
        }
        /*
          衝突Dirを保存
        */
        hit_dir_save();
        /*
          方向ブロック反射
        */
        vec_dir_reflect();
        /*
          衝突ポイントへ移動
        */
        pos_hit_point();
        timeout(S_VANISH);
        if (HasNextState())
        {
            GoNextState();
        }
    }
    /*
      DIR内の移動
    */
    void S_MOVE_DIR(bool bFirst)
    {
        if (bFirst)
        {
            SetNextState(null);
        }
        /*
          1ステップ移動
        */
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
    /*
      矢印ブロックに衝突
    */
    void S_HITYAJI(bool bFirst)
    {
        if (bFirst)
        {
            SetNextState(S_VANISH);
        }
        /*
          矢印を変更
        */
        yaji_change();
        timeout(S_VANISH);
        if (HasNextState())
        {
            GoNextState();
        }
    }
    /*
      ターゲットに衝突
    */
    void S_HITTARGET(bool bFirst)
    {
        if (bFirst)
        {
            SetNextState(S_VANISH);
        }
        /*
          ターゲット消滅
        */
        target_vanish();
        timeout(S_VANISH);
        if (HasNextState())
        {
            GoNextState();
        }
    }
    /*
      敵に衝突
    */
    void S_HITENEMY(bool bFirst)
    {
        if (bFirst)
        {
            SetNextState(S_MOVE);
        }
        /*
          ノーマル反射
        */
        vec_normal_reflect();
        /*
          衝突ポイントへ移動
        */
        pos_hit_point();
        /*
          敵消滅
        */
        enemy_vanish();
        timeout(S_VANISH);
        if (HasNextState())
        {
            GoNextState();
        }
    }
    /*
      ブロックに衝突
    */
    void S_HITBROCK(bool bFirst)
    {
        if (bFirst)
        {
            SetNextState(S_MOVE);
        }
        /*
          ノーマル反射
        */
        vec_normal_reflect();
        /*
          衝突ポイントへ移動
        */
        pos_hit_point();
        /*
          ブロック消滅
        */
        brock_vanish();
        timeout(S_VANISH);
        if (HasNextState())
        {
            GoNextState();
        }
    }
    /*
      障害物に衝突
    */
    void S_HITOBS(bool bFirst)
    {
        if (bFirst)
        {
            SetNextState(S_MOVE);
        }
        /*
          ノーマル反射
        */
        vec_normal_reflect();
        /*
          衝突ポイントへ移動
        */
        pos_hit_point();
        timeout(S_VANISH);
        if (HasNextState())
        {
            GoNextState();
        }
    }
    /*
      弾消滅
    */
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
    
}
