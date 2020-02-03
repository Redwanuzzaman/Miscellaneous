Dim speaks, speech
speaks="Biss Millah hir Rahh Manir Rahim"
set speech=CreateObject("sapi.spvoice")
with speech
       Set.voice = .getvoices.item(0)
       .Volume = 100
       .Rate = -4
   end with
speech.speak speaks