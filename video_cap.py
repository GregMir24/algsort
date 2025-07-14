import cv2


cap = cv2.VideoCapture(1)
count = 0
while True:
    ret, frame = cap.read()
    if not ret:
        print('error')
        break

    cv2.imshow('cam', frame)
    key = cv2.waitKey(1)
    if key == ord(' '):
        count += 1
        cv2.imwrite(f'Q:/ITMOPROJECT/img_collecting/img_collect_{count}.jpg', frame)
        print('saved')
        continue

    elif key == ord('e'):
        break

cap.release()
cv2.destroyAllWindows()