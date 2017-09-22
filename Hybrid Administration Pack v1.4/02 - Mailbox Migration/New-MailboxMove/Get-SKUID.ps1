#################################################################################################################
###                                                                                                           ###
###  	Script by Terry Munro -                                                                               ###
###     Technical Blog -               http://365admin.com.au                                                 ###
###     Webpage -                      https://www.linkedin.com/in/terry-munro/                               ###
###     TechNet Gallery Scripts -      http://tinyurl.com/TerryMunroTechNet                                   ###
###                                                                                                           ###
###     TechNet Download link -        https://gallery.technet.microsoft.com/Hybrid-Office-365-9b4570a5       ###
###                                                                                                           ###
###     Version 1.0 - 25/06/2017                                                                              ###
###                                                                                                           ###
###     Revision -                                                                                            ###
###               v1.0 - Initial script                                                                       ###
###                                                                                                           ###
#################################################################################################################

Get-MsolAccountSku | Format-Table AccountSkuId, SkuPartNumber