<?php

# This file is for remotely disabling the automated_students scripts for Arkansas Public Schools.
# The purpose of this script is if there is a reason to stop all imports because of unexpected
# consequences when the State upgrades/changes eSchool and Cognos.

$response->name = "Automated Students Remote Check"; #Required
$response->status = "OK"; #change to DISABLED

$response->version = "1.2"; #Optional

echo json_encode($response);

?>