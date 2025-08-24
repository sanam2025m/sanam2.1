document.getElementById("reportForm").addEventListener("submit", async function(e) {
  e.preventDefault();

  const submitBtn = this.querySelector("button[type='submit']");
  const responseBox = document.getElementById("response");
  responseBox.innerText = "";

  // تعطيل الزر لمنع التكرار
  if (submitBtn.disabled) return;
  submitBtn.disabled = true;

  // إظهار التنبيه الأول
  submitBtn.innerText = "🚨 جاري رفع المرفقات...";

  const fullName = document.getElementById("fullName").value;
  const incidentDate = document.getElementById("incidentDate").value;
  const startTime = document.getElementById("startTime").value;
  const location = document.getElementById("location").value;
  const uniqueKey = fullName + "_" + incidentDate + "_" + startTime + "_" + location;

  const formData = new FormData();
  formData.append("location", location);
  formData.append("fullName", fullName);
  formData.append("incidentDate", incidentDate);
  formData.append("startTime", startTime);
  formData.append("description", document.getElementById("description").value);
  formData.append("endTime", document.getElementById("endTime").value);
  formData.append("clientNotified", document.getElementById("clientNotified").value);
  formData.append("clientInstruction", document.getElementById("clientInstruction").value);
  formData.append("uniqueKey", uniqueKey);

  const files = document.getElementById("attachments").files;
  formData.append("fileCount", files.length);

  const readFileAsBase64 = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result.split(",")[1]);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  };

  try {
    for (let i = 0; i < files.length; i++) {
      const base64 = await readFileAsBase64(files[i]);
      formData.append("file" + i, base64);
      formData.append("name" + i, files[i].name);
      formData.append("type" + i, files[i].type);
    }

    // تغيير النص أثناء الإرسال
    submitBtn.innerText = "📤 جاري إرسال البلاغ...";

    const response = await fetch("https://script.google.com/macros/s/AKfycbzApOQ-dCNGAdqE_ROAb0KkmS0RSR3zGZrUajKMBeQ5AJGNRGS0hflEcIHhPMcvxQdN/exec", {
      method: "POST",
      body: formData
    });

    const result = await response.text();
    responseBox.innerText = result;
    responseBox.style.color = result.includes("✅") ? "green" : "red";
    document.getElementById("reportForm").reset();
  } catch (error) {
    responseBox.innerText = "❌ فشل في إرسال البيانات، تأكد من الاتصال.";
    responseBox.style.color = "red";
    console.error(error);
  }

  submitBtn.disabled = false;
  submitBtn.innerText = "إرسال البلاغ";
});

