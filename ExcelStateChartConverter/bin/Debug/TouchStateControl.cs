public enum TouchState
{
    S_NONE
    ,S_WAITTOUCH
    ,S_MOVE_PLAYERSYMBOL
    ,S_TOUCHED_HITBLOCK
    ,S_SET_HITBLOCK
    ,S_ROTATE_HITBLOCK
    ,S_DELETE_HITBLOCK

}
public partial class TouchStateControl {
    /*
        無指定
    */
    void S_NONE(bool bFist)
    {
        if (bFirst)
        {
            SetNextState(S_NONE);
        }
        GoNextState();
    }
    /*
        タッチ待ち
    */
    void S_WAITTOUCH(bool bFist)
    {
        if (bFirst)
        {
            SetNextState(null);
        }
        /*
            入力変化検知[タッチ入/解除・ シフト(指ズラシ)・ 同軸方向シフト・異軸方向シフト]
        */
        if (!(input_isChanged()))
        {
            return;
        }
        GoNextState();
    }
    /*
        プレイヤシンボル移動
    */
    void S_MOVE_PLAYERSYMBOL(bool bFist)
    {
        if (bFirst)
        {
            SetNextState(null);
            /*
                プレイヤ移動開始
            */
            ps_move_start();
        }
        /*
            入力変化検知
        */
        if (!(input_isChanged()))
        {
            return;
        }
        GoNextState();
    }
    /*
        ヒットブロックをタッチ時
    */
    void S_TOUCHED_HITBLOCK(bool bFist)
    {
        if (bFirst)
        {
            SetNextState(S_ROTATE_HITBLOCK);
            /*
                ヒットブロックを別種類へ変更
            */
            ps_alt_hitblock();
        }
        /*
            入力変化検知
        */
        if (!(input_is_changed()))
        {
            return;
        }
        GoNextState();
    }
    /*
        ヒットブロック設定
    */
    void S_SET_HITBLOCK(bool bFist)
    {
        if (bFirst)
        {
            SetNextState(S_ROTATE_HITBLOCK);
            /*
                hitblockを配置
                既に配置してあれば何もしない
                場所がない場合は配置なし
            */
            ps_set_hitblock();
        }
        /*
            入力変化検知
        */
        if (!(input_isChanged()))
        {
            return;
        }
        GoNextState();
    }
    /*
        ヒットブロックを回転する
    */
    void S_ROTATE_HITBLOCK(bool bFist)
    {
        if (bFirst)
        {
            SetNextState(null);
            /*
                hitblockを回転する
            */
            ps_rotate_hitblock();
        }
        /*
            入力変化検知
        */
        if (!(input_isChanged()))
        {
            return;
        }
        GoNextState();
    }
    /*
        ヒットブロックを削除
    */
    void S_DELETE_HITBLOCK(bool bFist)
    {
        if (bFirst)
        {
            SetNextState(S_WAITTOUCH);
            /*
                hitblockを削除
            */
            ps_delete_hitblock();
        }
        /*
            タッチが解除されるまで待つ
            ※解除後も0.1secほど待つように
        */
        if (!(input_isLeaveTouch()))
        {
            return;
        }
        GoNextState();
    }
    
}
