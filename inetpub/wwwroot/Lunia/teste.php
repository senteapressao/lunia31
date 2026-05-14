<?php
    $manager = new MongoDB\Driver\Manager(
        "mongodb://lunia:".urlencode('YGhF=yFCY3H=etBkgmg2$N%d$v=vB^7B')."@192.168.87.87:27017"
    );
    $query = new MongoDB\Driver\Query(array('Hash' => '249'));
    $cursor = $manager->executeQuery('lunia.items', $query);
    print_r($cursor->toArray());
?>