# ===================================================================================================
#                                  SECTION: Runtime Initialization
# ---------------------------------------------------------------------------------------------------
# Description:
#   Runs once every function is defined, before the GUI is built:
#     - loads the MailslotSender interop type (used by the lane/scale command functions)
# ===================================================================================================

Ensure_MailslotSender_Type
