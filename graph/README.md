グラフを作成する時に使うマクロ.

- `make_edge.vba`
    - `MakeEdgeUndirected1ToN`
        - 一つ目に選択した頂点から二つ目以降に選択した頂点全てに無向辺を張る.
    - `MakeEdgeDirected1ToN`
        - 一つ目に選択した頂点から二つ目以降に選択した頂点全てに有向辺を張る.
- `select_only.vba`
    - `SelectOnlyEdges`
        - 現在選択中のオブジェクトのうち, 辺だけを選択した状態にする.
    - `SelectOnlyOvals`
        - 現在選択中のオブジェクトのうち, 楕円(円)だけを選択した状態にする.
    - `SelectOnlyDodecagons`
        - 現在選択中のオブジェクトのうち, 十角形だけを選択した状態にする.
    - `SelectOnlyTextBoxes`
        - 現在選択中のオブジェクトのうち, テキストボックスだけを選択した状態にする.
- `assign_number.vba`
    - `AssignNumbers`
        - テキストフレームに選択順で番号を振る.
        - 1から始まる連続した番号を振る.
        - 最初に選択した図形に既に番号が振られている場合, その番号から連続した番号を振る.
        - 例: 最初に選択した図形に5が振られている場合, 次に選択した図形には6が振られる.
