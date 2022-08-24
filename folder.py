import os

resume_folder = r"C:\Users\Admin\PycharmProjects\Reliable\Resume_folder"
resume = os.listdir(resume_folder)
print("\n", resume)

for i in range(len(resume)):
    resumes_path = os.path.join(resume_folder, resume[i])
    os.startfile(resumes_path)
    print(resumes_path)


