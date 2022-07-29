echo "Input Video Path"
read vidURL
echo "Input Audio Path"
read audURL

vidTS = "${vidURL}\all.ts"
vidMP4 = "${vidURL}\all.mp4"
audTS = "${audURL}\all.ts"
audMP4 = "${audURL}\all.mp4"

ffmpeg -i $vidTS -c copy $vidMP4
echo "Video TS converted to MP4"
ffmpeg -i $audTS -c copy $audMP4
echo "Audio TS converted to MP4"

ffmpeg -i $vidMP4 -i $audMP4 -c copy output.mp4
